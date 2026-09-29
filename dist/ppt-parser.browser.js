/**
 * @fefeding/ppt-parser v1.0.12
 * PPTX文件解析与序列化核心库，纯TS编写，支持解析PPTX为JSON结构、JSON序列化为标准PPTX文件，无框架依赖
 * MIT License
 */
var commonjsGlobal = typeof globalThis !== 'undefined' ? globalThis : typeof window !== 'undefined' ? window : typeof global !== 'undefined' ? global : typeof self !== 'undefined' ? self : {};

function getDefaultExportFromCjs (x) {
	return x && x.__esModule && Object.prototype.hasOwnProperty.call(x, 'default') ? x['default'] : x;
}

function commonjsRequire(path) {
	throw new Error('Could not dynamically require "' + path + '". Please configure the dynamicRequireTargets or/and ignoreDynamicRequires option of @rollup/plugin-commonjs appropriately for this require call to work.');
}

var jszip_min = {exports: {}};

/*!

JSZip v3.10.2 - A JavaScript class for generating and reading zip files
<http://stuartk.com/jszip>

(c) 2009-2016 Stuart Knightley <stuart [at] stuartk.com>
Dual licenced under the MIT license or GPLv3. See https://raw.github.com/Stuk/jszip/main/LICENSE.markdown.

JSZip uses the library pako released under the MIT license :
https://github.com/nodeca/pako/blob/main/LICENSE
*/

(function (module, exports) {
	!function(e){module.exports=e();}(function(){return function s(a,o,h){function u(r,e){if(!o[r]){if(!a[r]){var t="function"==typeof commonjsRequire&&commonjsRequire;if(!e&&t)return t(r,true);if(l)return l(r,true);var n=new Error("Cannot find module '"+r+"'");throw n.code="MODULE_NOT_FOUND",n}var i=o[r]={exports:{}};a[r][0].call(i.exports,function(e){var t=a[r][1][e];return u(t||e)},i,i.exports,s,a,o,h);}return o[r].exports}for(var l="function"==typeof commonjsRequire&&commonjsRequire,e=0;e<h.length;e++)u(h[e]);return u}({1:[function(e,t,r){var d=e("./utils"),c=e("./support"),p="ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz0123456789+/=";r.encode=function(e){for(var t,r,n,i,s,a,o,h=[],u=0,l=e.length,f=l,c="string"!==d.getTypeOf(e);u<e.length;)f=l-u,n=c?(t=e[u++],r=u<l?e[u++]:0,u<l?e[u++]:0):(t=e.charCodeAt(u++),r=u<l?e.charCodeAt(u++):0,u<l?e.charCodeAt(u++):0),i=t>>2,s=(3&t)<<4|r>>4,a=1<f?(15&r)<<2|n>>6:64,o=2<f?63&n:64,h.push(p.charAt(i)+p.charAt(s)+p.charAt(a)+p.charAt(o));return h.join("")},r.decode=function(e){var t,r,n,i,s,a,o=0,h=0,u="data:";if(e.substr(0,u.length)===u)throw new Error("Invalid base64 input, it looks like a data url.");var l,f=3*(e=e.replace(/[^A-Za-z0-9+/=]/g,"")).length/4;if(e.charAt(e.length-1)===p.charAt(64)&&f--,e.charAt(e.length-2)===p.charAt(64)&&f--,f%1!=0)throw new Error("Invalid base64 input, bad content length.");for(l=c.uint8array?new Uint8Array(0|f):new Array(0|f);o<e.length;)t=p.indexOf(e.charAt(o++))<<2|(i=p.indexOf(e.charAt(o++)))>>4,r=(15&i)<<4|(s=p.indexOf(e.charAt(o++)))>>2,n=(3&s)<<6|(a=p.indexOf(e.charAt(o++))),l[h++]=t,64!==s&&(l[h++]=r),64!==a&&(l[h++]=n);return l};},{"./support":30,"./utils":32}],2:[function(e,t,r){var n=e("./external"),i=e("./stream/DataWorker"),s=e("./stream/Crc32Probe"),a=e("./stream/DataLengthProbe");function o(e,t,r,n,i){this.compressedSize=e,this.uncompressedSize=t,this.crc32=r,this.compression=n,this.compressedContent=i;}o.prototype={getContentWorker:function(){var e=new i(n.Promise.resolve(this.compressedContent)).pipe(this.compression.uncompressWorker()).pipe(new a("data_length")),t=this;return e.on("end",function(){if(this.streamInfo.data_length!==t.uncompressedSize)throw new Error("Bug : uncompressed data size mismatch")}),e},getCompressedWorker:function(){return new i(n.Promise.resolve(this.compressedContent)).withStreamInfo("compressedSize",this.compressedSize).withStreamInfo("uncompressedSize",this.uncompressedSize).withStreamInfo("crc32",this.crc32).withStreamInfo("compression",this.compression)}},o.createWorkerFrom=function(e,t,r){return e.pipe(new s).pipe(new a("uncompressedSize")).pipe(t.compressWorker(r)).pipe(new a("compressedSize")).withStreamInfo("compression",t)},t.exports=o;},{"./external":6,"./stream/Crc32Probe":25,"./stream/DataLengthProbe":26,"./stream/DataWorker":27}],3:[function(e,t,r){var n=e("./stream/GenericWorker");r.STORE={magic:"\0\0",compressWorker:function(){return new n("STORE compression")},uncompressWorker:function(){return new n("STORE decompression")}},r.DEFLATE=e("./flate");},{"./flate":7,"./stream/GenericWorker":28}],4:[function(e,t,r){var n=e("./utils");var o=function(){for(var e,t=[],r=0;r<256;r++){e=r;for(var n=0;n<8;n++)e=1&e?3988292384^e>>>1:e>>>1;t[r]=e;}return t}();t.exports=function(e,t){return void 0!==e&&e.length?"string"!==n.getTypeOf(e)?function(e,t,r,n){var i=o,s=n+r;e^=-1;for(var a=n;a<s;a++)e=e>>>8^i[255&(e^t[a])];return  -1^e}(0|t,e,e.length,0):function(e,t,r,n){var i=o,s=n+r;e^=-1;for(var a=n;a<s;a++)e=e>>>8^i[255&(e^t.charCodeAt(a))];return  -1^e}(0|t,e,e.length,0):0};},{"./utils":32}],5:[function(e,t,r){r.base64=false,r.binary=false,r.dir=false,r.createFolders=true,r.date=null,r.compression=null,r.compressionOptions=null,r.comment=null,r.unixPermissions=null,r.dosPermissions=null;},{}],6:[function(e,t,r){var n=null;n="undefined"!=typeof Promise?Promise:e("lie"),t.exports={Promise:n};},{lie:37}],7:[function(e,t,r){var n="undefined"!=typeof Uint8Array&&"undefined"!=typeof Uint16Array&&"undefined"!=typeof Uint32Array,i=e("pako"),s=e("./utils"),a=e("./stream/GenericWorker"),o=n?"uint8array":"array";function h(e,t){a.call(this,"FlateWorker/"+e),this._pako=null,this._pakoAction=e,this._pakoOptions=t,this.meta={};}r.magic="\b\0",s.inherits(h,a),h.prototype.processChunk=function(e){this.meta=e.meta,null===this._pako&&this._createPako(),this._pako.push(s.transformTo(o,e.data),false);},h.prototype.flush=function(){a.prototype.flush.call(this),null===this._pako&&this._createPako(),this._pako.push([],true);},h.prototype.cleanUp=function(){a.prototype.cleanUp.call(this),this._pako=null;},h.prototype._createPako=function(){this._pako=new i[this._pakoAction]({raw:true,level:this._pakoOptions.level||-1});var t=this;this._pako.onData=function(e){t.push({data:e,meta:t.meta});};},r.compressWorker=function(e){return new h("Deflate",e)},r.uncompressWorker=function(){return new h("Inflate",{})};},{"./stream/GenericWorker":28,"./utils":32,pako:38}],8:[function(e,t,r){function A(e,t){var r,n="";for(r=0;r<t;r++)n+=String.fromCharCode(255&e),e>>>=8;return n}function n(e,t,r,n,i,s){var a,o,h=e.file,u=e.compression,l=s!==O.utf8encode,f=I.transformTo("string",s(h.name)),c=I.transformTo("string",O.utf8encode(h.name)),d=h.comment,p=I.transformTo("string",s(d)),m=I.transformTo("string",O.utf8encode(d)),_=c.length!==h.name.length,g=m.length!==d.length,b="",v="",y="",w=h.dir,k=h.date,x={crc32:0,compressedSize:0,uncompressedSize:0};t&&!r||(x.crc32=e.crc32,x.compressedSize=e.compressedSize,x.uncompressedSize=e.uncompressedSize);var S=0;t&&(S|=8),l||!_&&!g||(S|=2048);var z=0,C=0;w&&(z|=16),"UNIX"===i?(C=798,z|=function(e,t){var r=e;return e||(r=t?16893:33204),(65535&r)<<16}(h.unixPermissions,w)):(C=20,z|=function(e){return 63&(e||0)}(h.dosPermissions)),a=k.getUTCHours(),a<<=6,a|=k.getUTCMinutes(),a<<=5,a|=k.getUTCSeconds()/2,o=k.getUTCFullYear()-1980,o<<=4,o|=k.getUTCMonth()+1,o<<=5,o|=k.getUTCDate(),_&&(v=A(1,1)+A(B(f),4)+c,b+="up"+A(v.length,2)+v),g&&(y=A(1,1)+A(B(p),4)+m,b+="uc"+A(y.length,2)+y);var E="";return E+="\n\0",E+=A(S,2),E+=u.magic,E+=A(a,2),E+=A(o,2),E+=A(x.crc32,4),E+=A(x.compressedSize,4),E+=A(x.uncompressedSize,4),E+=A(f.length,2),E+=A(b.length,2),{fileRecord:R.LOCAL_FILE_HEADER+E+f+b,dirRecord:R.CENTRAL_FILE_HEADER+A(C,2)+E+A(p.length,2)+"\0\0\0\0"+A(z,4)+A(n,4)+f+b+p}}var I=e("../utils"),i=e("../stream/GenericWorker"),O=e("../utf8"),B=e("../crc32"),R=e("../signature");function s(e,t,r,n){i.call(this,"ZipFileWorker"),this.bytesWritten=0,this.zipComment=t,this.zipPlatform=r,this.encodeFileName=n,this.streamFiles=e,this.accumulate=false,this.contentBuffer=[],this.dirRecords=[],this.currentSourceOffset=0,this.entriesCount=0,this.currentFile=null,this._sources=[];}I.inherits(s,i),s.prototype.push=function(e){var t=e.meta.percent||0,r=this.entriesCount,n=this._sources.length;this.accumulate?this.contentBuffer.push(e):(this.bytesWritten+=e.data.length,i.prototype.push.call(this,{data:e.data,meta:{currentFile:this.currentFile,percent:r?(t+100*(r-n-1))/r:100}}));},s.prototype.openedSource=function(e){this.currentSourceOffset=this.bytesWritten,this.currentFile=e.file.name;var t=this.streamFiles&&!e.file.dir;if(t){var r=n(e,t,false,this.currentSourceOffset,this.zipPlatform,this.encodeFileName);this.push({data:r.fileRecord,meta:{percent:0}});}else this.accumulate=true;},s.prototype.closedSource=function(e){this.accumulate=false;var t=this.streamFiles&&!e.file.dir,r=n(e,t,true,this.currentSourceOffset,this.zipPlatform,this.encodeFileName);if(this.dirRecords.push(r.dirRecord),t)this.push({data:function(e){return R.DATA_DESCRIPTOR+A(e.crc32,4)+A(e.compressedSize,4)+A(e.uncompressedSize,4)}(e),meta:{percent:100}});else for(this.push({data:r.fileRecord,meta:{percent:0}});this.contentBuffer.length;)this.push(this.contentBuffer.shift());this.currentFile=null;},s.prototype.flush=function(){for(var e=this.bytesWritten,t=0;t<this.dirRecords.length;t++)this.push({data:this.dirRecords[t],meta:{percent:100}});var r=this.bytesWritten-e,n=function(e,t,r,n,i){var s=I.transformTo("string",i(n));return R.CENTRAL_DIRECTORY_END+"\0\0\0\0"+A(e,2)+A(e,2)+A(t,4)+A(r,4)+A(s.length,2)+s}(this.dirRecords.length,r,e,this.zipComment,this.encodeFileName);this.push({data:n,meta:{percent:100}});},s.prototype.prepareNextSource=function(){this.previous=this._sources.shift(),this.openedSource(this.previous.streamInfo),this.isPaused?this.previous.pause():this.previous.resume();},s.prototype.registerPrevious=function(e){this._sources.push(e);var t=this;return e.on("data",function(e){t.processChunk(e);}),e.on("end",function(){t.closedSource(t.previous.streamInfo),t._sources.length?t.prepareNextSource():t.end();}),e.on("error",function(e){t.error(e);}),this},s.prototype.resume=function(){return !!i.prototype.resume.call(this)&&(!this.previous&&this._sources.length?(this.prepareNextSource(),true):this.previous||this._sources.length||this.generatedError?void 0:(this.end(),true))},s.prototype.error=function(e){var t=this._sources;if(!i.prototype.error.call(this,e))return  false;for(var r=0;r<t.length;r++)try{t[r].error(e);}catch(e){}return  true},s.prototype.lock=function(){i.prototype.lock.call(this);for(var e=this._sources,t=0;t<e.length;t++)e[t].lock();},t.exports=s;},{"../crc32":4,"../signature":23,"../stream/GenericWorker":28,"../utf8":31,"../utils":32}],9:[function(e,t,r){var u=e("../compressions"),n=e("./ZipFileWorker");r.generateWorker=function(e,a,t){var o=new n(a.streamFiles,t,a.platform,a.encodeFileName),h=0;try{e.forEach(function(e,t){h++;var r=function(e,t){var r=e||t,n=u[r];if(!n)throw new Error(r+" is not a valid compression method !");return n}(t.options.compression,a.compression),n=t.options.compressionOptions||a.compressionOptions||{},i=t.dir,s=t.date;t._compressWorker(r,n).withStreamInfo("file",{name:e,dir:i,date:s,comment:t.comment||"",unixPermissions:t.unixPermissions,dosPermissions:t.dosPermissions}).pipe(o);}),o.entriesCount=h;}catch(e){o.error(e);}return o};},{"../compressions":3,"./ZipFileWorker":8}],10:[function(e,t,r){function n(){if(!(this instanceof n))return new n;if(arguments.length)throw new Error("The constructor with parameters has been removed in JSZip 3.0, please check the upgrade guide.");this.files=Object.create(null),this.comment=null,this.root="",this.clone=function(){var e=new n;for(var t in this)"function"!=typeof this[t]&&(e[t]=this[t]);return e};}(n.prototype=e("./object")).loadAsync=e("./load"),n.support=e("./support"),n.defaults=e("./defaults"),n.version="3.10.2",n.loadAsync=function(e,t){return (new n).loadAsync(e,t)},n.external=e("./external"),t.exports=n;},{"./defaults":5,"./external":6,"./load":11,"./object":15,"./support":30}],11:[function(e,t,r){var u=e("./utils"),i=e("./external"),n=e("./utf8"),s=e("./zipEntries"),a=e("./stream/Crc32Probe"),l=e("./nodejsUtils");function f(n){return new i.Promise(function(e,t){var r=n.decompressed.getContentWorker().pipe(new a);r.on("error",function(e){t(e);}).on("end",function(){r.streamInfo.crc32!==n.decompressed.crc32?t(new Error("Corrupted zip : CRC32 mismatch")):e();}).resume();})}t.exports=function(e,o){var h=this;return o=u.extend(o||{},{base64:false,checkCRC32:false,optimizedBinaryString:false,createFolders:false,decodeFileName:n.utf8decode}),l.isNode&&l.isStream(e)?i.Promise.reject(new Error("JSZip can't accept a stream when loading a zip file.")):u.prepareContent("the loaded zip file",e,true,o.optimizedBinaryString,o.base64).then(function(e){var t=new s(o);return t.load(e),t}).then(function(e){var t=[i.Promise.resolve(e)],r=e.files;if(o.checkCRC32)for(var n=0;n<r.length;n++)t.push(f(r[n]));return i.Promise.all(t)}).then(function(e){for(var t=e.shift(),r=t.files,n=0;n<r.length;n++){var i=r[n],s=i.fileNameStr,a=u.resolve(i.fileNameStr);h.file(a,i.decompressed,{binary:true,optimizedBinaryString:true,date:i.date,dir:i.dir,comment:i.fileCommentStr.length?i.fileCommentStr:null,unixPermissions:i.unixPermissions,dosPermissions:i.dosPermissions,createFolders:o.createFolders}),i.dir||(h.file(a).unsafeOriginalName=s);}return t.zipComment.length&&(h.comment=t.zipComment),h})};},{"./external":6,"./nodejsUtils":14,"./stream/Crc32Probe":25,"./utf8":31,"./utils":32,"./zipEntries":33}],12:[function(e,t,r){var n=e("../utils"),i=e("../stream/GenericWorker");function s(e,t){i.call(this,"Nodejs stream input adapter for "+e),this._upstreamEnded=false,this._bindStream(t);}n.inherits(s,i),s.prototype._bindStream=function(e){var t=this;(this._stream=e).pause(),e.on("data",function(e){t.push({data:e,meta:{percent:0}});}).on("error",function(e){t.isPaused?this.generatedError=e:t.error(e);}).on("end",function(){t.isPaused?t._upstreamEnded=true:t.end();});},s.prototype.pause=function(){return !!i.prototype.pause.call(this)&&(this._stream.pause(),true)},s.prototype.resume=function(){return !!i.prototype.resume.call(this)&&(this._upstreamEnded?this.end():this._stream.resume(),true)},t.exports=s;},{"../stream/GenericWorker":28,"../utils":32}],13:[function(e,t,r){var i=e("readable-stream").Readable;function n(e,t,r){i.call(this,t),this._helper=e;var n=this;e.on("data",function(e,t){n.push(e)||n._helper.pause(),r&&r(t);}).on("error",function(e){n.emit("error",e);}).on("end",function(){n.push(null);});}e("../utils").inherits(n,i),n.prototype._read=function(){this._helper.resume();},t.exports=n;},{"../utils":32,"readable-stream":16}],14:[function(e,t,r){t.exports={isNode:"undefined"!=typeof Buffer,newBufferFrom:function(e,t){if(Buffer.from&&Buffer.from!==Uint8Array.from)return Buffer.from(e,t);if("number"==typeof e)throw new Error('The "data" argument must not be a number');return new Buffer(e,t)},allocBuffer:function(e){if(Buffer.alloc)return Buffer.alloc(e);var t=new Buffer(e);return t.fill(0),t},isBuffer:function(e){return Buffer.isBuffer(e)},isStream:function(e){return e&&"function"==typeof e.on&&"function"==typeof e.pause&&"function"==typeof e.resume}};},{}],15:[function(e,t,r){function s(e,t,r){var n,i=u.getTypeOf(t),s=u.extend(r||{},f);s.date=s.date||new Date,null!==s.compression&&(s.compression=s.compression.toUpperCase()),"string"==typeof s.unixPermissions&&(s.unixPermissions=parseInt(s.unixPermissions,8)),s.unixPermissions&&16384&s.unixPermissions&&(s.dir=true),s.dosPermissions&&16&s.dosPermissions&&(s.dir=true),s.dir&&(e=g(e)),s.createFolders&&(n=_(e))&&b.call(this,n,true);var a="string"===i&&false===s.binary&&false===s.base64;r&&void 0!==r.binary||(s.binary=!a),(t instanceof c&&0===t.uncompressedSize||s.dir||!t||0===t.length)&&(s.base64=false,s.binary=true,t="",s.compression="STORE",i="string");var o=null;o=t instanceof c||t instanceof l?t:p.isNode&&p.isStream(t)?new m(e,t):u.prepareContent(e,t,s.binary,s.optimizedBinaryString,s.base64);var h=new d(e,o,s);this.files[e]=h;}var i=e("./utf8"),u=e("./utils"),l=e("./stream/GenericWorker"),a=e("./stream/StreamHelper"),f=e("./defaults"),c=e("./compressedObject"),d=e("./zipObject"),o=e("./generate"),p=e("./nodejsUtils"),m=e("./nodejs/NodejsStreamInputAdapter"),_=function(e){"/"===e.slice(-1)&&(e=e.substring(0,e.length-1));var t=e.lastIndexOf("/");return 0<t?e.substring(0,t):""},g=function(e){return "/"!==e.slice(-1)&&(e+="/"),e},b=function(e,t){return t=void 0!==t?t:f.createFolders,e=g(e),this.files[e]||s.call(this,e,null,{dir:true,createFolders:t}),this.files[e]};function h(e){return "[object RegExp]"===Object.prototype.toString.call(e)}var n={load:function(){throw new Error("This method has been removed in JSZip 3.0, please check the upgrade guide.")},forEach:function(e){var t,r,n;for(t in this.files)n=this.files[t],(r=t.slice(this.root.length,t.length))&&t.slice(0,this.root.length)===this.root&&e(r,n);},filter:function(r){var n=[];return this.forEach(function(e,t){r(e,t)&&n.push(t);}),n},file:function(e,t,r){if(1!==arguments.length)return e=this.root+e,s.call(this,e,t,r),this;if(h(e)){var n=e;return this.filter(function(e,t){return !t.dir&&n.test(e)})}var i=this.files[this.root+e];return i&&!i.dir?i:null},folder:function(r){if(!r)return this;if(h(r))return this.filter(function(e,t){return t.dir&&r.test(e)});var e=this.root+r,t=b.call(this,e),n=this.clone();return n.root=t.name,n},remove:function(r){r=this.root+r;var e=this.files[r];if(e||("/"!==r.slice(-1)&&(r+="/"),e=this.files[r]),e&&!e.dir)delete this.files[r];else for(var t=this.filter(function(e,t){return t.name.slice(0,r.length)===r}),n=0;n<t.length;n++)delete this.files[t[n].name];return this},generate:function(){throw new Error("This method has been removed in JSZip 3.0, please check the upgrade guide.")},generateInternalStream:function(e){var t,r={};try{if((r=u.extend(e||{},{streamFiles:!1,compression:"STORE",compressionOptions:null,type:"",platform:"DOS",comment:null,mimeType:"application/zip",encodeFileName:i.utf8encode})).type=r.type.toLowerCase(),r.compression=r.compression.toUpperCase(),"binarystring"===r.type&&(r.type="string"),!r.type)throw new Error("No output type specified.");u.checkSupport(r.type),"darwin"!==r.platform&&"freebsd"!==r.platform&&"linux"!==r.platform&&"sunos"!==r.platform||(r.platform="UNIX"),"win32"===r.platform&&(r.platform="DOS");var n=r.comment||this.comment||"";t=o.generateWorker(this,r,n);}catch(e){(t=new l("error")).error(e);}return new a(t,r.type||"string",r.mimeType)},generateAsync:function(e,t){return this.generateInternalStream(e).accumulate(t)},generateNodeStream:function(e,t){return (e=e||{}).type||(e.type="nodebuffer"),this.generateInternalStream(e).toNodejsStream(t)}};t.exports=n;},{"./compressedObject":2,"./defaults":5,"./generate":9,"./nodejs/NodejsStreamInputAdapter":12,"./nodejsUtils":14,"./stream/GenericWorker":28,"./stream/StreamHelper":29,"./utf8":31,"./utils":32,"./zipObject":35}],16:[function(e,t,r){t.exports=e("stream");},{stream:void 0}],17:[function(e,t,r){var n=e("./DataReader");function i(e){n.call(this,e);for(var t=0;t<this.data.length;t++)e[t]=255&e[t];}e("../utils").inherits(i,n),i.prototype.byteAt=function(e){return this.data[this.zero+e]},i.prototype.lastIndexOfSignature=function(e){for(var t=e.charCodeAt(0),r=e.charCodeAt(1),n=e.charCodeAt(2),i=e.charCodeAt(3),s=this.length-4;0<=s;--s)if(this.data[s]===t&&this.data[s+1]===r&&this.data[s+2]===n&&this.data[s+3]===i)return s-this.zero;return  -1},i.prototype.readAndCheckSignature=function(e){var t=e.charCodeAt(0),r=e.charCodeAt(1),n=e.charCodeAt(2),i=e.charCodeAt(3),s=this.readData(4);return t===s[0]&&r===s[1]&&n===s[2]&&i===s[3]},i.prototype.readData=function(e){if(this.checkOffset(e),0===e)return [];var t=this.data.slice(this.zero+this.index,this.zero+this.index+e);return this.index+=e,t},t.exports=i;},{"../utils":32,"./DataReader":18}],18:[function(e,t,r){var n=e("../utils");function i(e){this.data=e,this.length=e.length,this.index=0,this.zero=0;}i.prototype={checkOffset:function(e){this.checkIndex(this.index+e);},checkIndex:function(e){if(this.length<this.zero+e||e<0)throw new Error("End of data reached (data length = "+this.length+", asked index = "+e+"). Corrupted zip ?")},setIndex:function(e){this.checkIndex(e),this.index=e;},skip:function(e){this.setIndex(this.index+e);},byteAt:function(){},readInt:function(e){var t,r=0;for(this.checkOffset(e),t=this.index+e-1;t>=this.index;t--)r=(r<<8)+this.byteAt(t);return this.index+=e,r},readString:function(e){return n.transformTo("string",this.readData(e))},readData:function(){},lastIndexOfSignature:function(){},readAndCheckSignature:function(){},readDate:function(){var e=this.readInt(4);return new Date(Date.UTC(1980+(e>>25&127),(e>>21&15)-1,e>>16&31,e>>11&31,e>>5&63,(31&e)<<1))}},t.exports=i;},{"../utils":32}],19:[function(e,t,r){var n=e("./Uint8ArrayReader");function i(e){n.call(this,e);}e("../utils").inherits(i,n),i.prototype.readData=function(e){this.checkOffset(e);var t=this.data.slice(this.zero+this.index,this.zero+this.index+e);return this.index+=e,t},t.exports=i;},{"../utils":32,"./Uint8ArrayReader":21}],20:[function(e,t,r){var n=e("./DataReader");function i(e){n.call(this,e);}e("../utils").inherits(i,n),i.prototype.byteAt=function(e){return this.data.charCodeAt(this.zero+e)},i.prototype.lastIndexOfSignature=function(e){return this.data.lastIndexOf(e)-this.zero},i.prototype.readAndCheckSignature=function(e){return e===this.readData(4)},i.prototype.readData=function(e){this.checkOffset(e);var t=this.data.slice(this.zero+this.index,this.zero+this.index+e);return this.index+=e,t},t.exports=i;},{"../utils":32,"./DataReader":18}],21:[function(e,t,r){var n=e("./ArrayReader");function i(e){n.call(this,e);}e("../utils").inherits(i,n),i.prototype.readData=function(e){if(this.checkOffset(e),0===e)return new Uint8Array(0);var t=this.data.subarray(this.zero+this.index,this.zero+this.index+e);return this.index+=e,t},t.exports=i;},{"../utils":32,"./ArrayReader":17}],22:[function(e,t,r){var n=e("../utils"),i=e("../support"),s=e("./ArrayReader"),a=e("./StringReader"),o=e("./NodeBufferReader"),h=e("./Uint8ArrayReader");t.exports=function(e){var t=n.getTypeOf(e);return n.checkSupport(t),"string"!==t||i.uint8array?"nodebuffer"===t?new o(e):i.uint8array?new h(n.transformTo("uint8array",e)):new s(n.transformTo("array",e)):new a(e)};},{"../support":30,"../utils":32,"./ArrayReader":17,"./NodeBufferReader":19,"./StringReader":20,"./Uint8ArrayReader":21}],23:[function(e,t,r){r.LOCAL_FILE_HEADER="PK",r.CENTRAL_FILE_HEADER="PK",r.CENTRAL_DIRECTORY_END="PK",r.ZIP64_CENTRAL_DIRECTORY_LOCATOR="PK",r.ZIP64_CENTRAL_DIRECTORY_END="PK",r.DATA_DESCRIPTOR="PK\b";},{}],24:[function(e,t,r){var n=e("./GenericWorker"),i=e("../utils");function s(e){n.call(this,"ConvertWorker to "+e),this.destType=e;}i.inherits(s,n),s.prototype.processChunk=function(e){this.push({data:i.transformTo(this.destType,e.data),meta:e.meta});},t.exports=s;},{"../utils":32,"./GenericWorker":28}],25:[function(e,t,r){var n=e("./GenericWorker"),i=e("../crc32");function s(){n.call(this,"Crc32Probe"),this.withStreamInfo("crc32",0);}e("../utils").inherits(s,n),s.prototype.processChunk=function(e){this.streamInfo.crc32=i(e.data,this.streamInfo.crc32||0),this.push(e);},t.exports=s;},{"../crc32":4,"../utils":32,"./GenericWorker":28}],26:[function(e,t,r){var n=e("../utils"),i=e("./GenericWorker");function s(e){i.call(this,"DataLengthProbe for "+e),this.propName=e,this.withStreamInfo(e,0);}n.inherits(s,i),s.prototype.processChunk=function(e){if(e){var t=this.streamInfo[this.propName]||0;this.streamInfo[this.propName]=t+e.data.length;}i.prototype.processChunk.call(this,e);},t.exports=s;},{"../utils":32,"./GenericWorker":28}],27:[function(e,t,r){var n=e("../utils"),i=e("./GenericWorker");function s(e){i.call(this,"DataWorker");var t=this;this.dataIsReady=false,this.index=0,this.max=0,this.data=null,this.type="",this._tickScheduled=false,e.then(function(e){t.dataIsReady=true,t.data=e,t.max=e&&e.length||0,t.type=n.getTypeOf(e),t.isPaused||t._tickAndRepeat();},function(e){t.error(e);});}n.inherits(s,i),s.prototype.cleanUp=function(){i.prototype.cleanUp.call(this),this.data=null;},s.prototype.resume=function(){return !!i.prototype.resume.call(this)&&(!this._tickScheduled&&this.dataIsReady&&(this._tickScheduled=true,n.delay(this._tickAndRepeat,[],this)),true)},s.prototype._tickAndRepeat=function(){this._tickScheduled=false,this.isPaused||this.isFinished||(this._tick(),this.isFinished||(n.delay(this._tickAndRepeat,[],this),this._tickScheduled=true));},s.prototype._tick=function(){if(this.isPaused||this.isFinished)return  false;var e=null,t=Math.min(this.max,this.index+16384);if(this.index>=this.max)return this.end();switch(this.type){case "string":e=this.data.substring(this.index,t);break;case "uint8array":e=this.data.subarray(this.index,t);break;case "array":case "nodebuffer":e=this.data.slice(this.index,t);}return this.index=t,this.push({data:e,meta:{percent:this.max?this.index/this.max*100:0}})},t.exports=s;},{"../utils":32,"./GenericWorker":28}],28:[function(e,t,r){function n(e){this.name=e||"default",this.streamInfo={},this.generatedError=null,this.extraStreamInfo={},this.isPaused=true,this.isFinished=false,this.isLocked=false,this._listeners={data:[],end:[],error:[]},this.previous=null;}n.prototype={push:function(e){this.emit("data",e);},end:function(){if(this.isFinished)return  false;this.flush();try{this.emit("end"),this.cleanUp(),this.isFinished=!0;}catch(e){this.emit("error",e);}return  true},error:function(e){return !this.isFinished&&(this.isPaused?this.generatedError=e:(this.isFinished=true,this.emit("error",e),this.previous&&this.previous.error(e),this.cleanUp()),true)},on:function(e,t){return this._listeners[e].push(t),this},cleanUp:function(){this.streamInfo=this.generatedError=this.extraStreamInfo=null,this._listeners=[];},emit:function(e,t){if(this._listeners[e])for(var r=0;r<this._listeners[e].length;r++)this._listeners[e][r].call(this,t);},pipe:function(e){return e.registerPrevious(this)},registerPrevious:function(e){if(this.isLocked)throw new Error("The stream '"+this+"' has already been used.");this.streamInfo=e.streamInfo,this.mergeStreamInfo(),this.previous=e;var t=this;return e.on("data",function(e){t.processChunk(e);}),e.on("end",function(){t.end();}),e.on("error",function(e){t.error(e);}),this},pause:function(){return !this.isPaused&&!this.isFinished&&(this.isPaused=true,this.previous&&this.previous.pause(),true)},resume:function(){if(!this.isPaused||this.isFinished)return  false;var e=this.isPaused=false;return this.generatedError&&(this.error(this.generatedError),e=true),this.previous&&this.previous.resume(),!e},flush:function(){},processChunk:function(e){this.push(e);},withStreamInfo:function(e,t){return this.extraStreamInfo[e]=t,this.mergeStreamInfo(),this},mergeStreamInfo:function(){for(var e in this.extraStreamInfo)Object.prototype.hasOwnProperty.call(this.extraStreamInfo,e)&&(this.streamInfo[e]=this.extraStreamInfo[e]);},lock:function(){if(this.isLocked)throw new Error("The stream '"+this+"' has already been used.");this.isLocked=true,this.previous&&this.previous.lock();},toString:function(){var e="Worker "+this.name;return this.previous?this.previous+" -> "+e:e}},t.exports=n;},{}],29:[function(e,t,r){var h=e("../utils"),i=e("./ConvertWorker"),s=e("./GenericWorker"),u=e("../base64"),n=e("../support"),a=e("../external"),o=null;if(n.nodestream)try{o=e("../nodejs/NodejsStreamOutputAdapter");}catch(e){}function l(e,o){return new a.Promise(function(t,r){var n=[],i=e._internalType,s=e._outputType,a=e._mimeType;e.on("data",function(e,t){n.push(e),o&&o(t);}).on("error",function(e){n=[],r(e);}).on("end",function(){try{var e=function(e,t,r){switch(e){case "blob":return h.newBlob(h.transformTo("arraybuffer",t),r);case "base64":return u.encode(t);default:return h.transformTo(e,t)}}(s,function(e,t){var r,n=0,i=null,s=0;for(r=0;r<t.length;r++)s+=t[r].length;switch(e){case "string":return t.join("");case "array":return Array.prototype.concat.apply([],t);case "uint8array":for(i=new Uint8Array(s),r=0;r<t.length;r++)i.set(t[r],n),n+=t[r].length;return i;case "nodebuffer":return Buffer.concat(t);default:throw new Error("concat : unsupported type '"+e+"'")}}(i,n),a);t(e);}catch(e){r(e);}n=[];}).resume();})}function f(e,t,r){var n=t;switch(t){case "blob":case "arraybuffer":n="uint8array";break;case "base64":n="string";}try{this._internalType=n,this._outputType=t,this._mimeType=r,h.checkSupport(n),this._worker=e.pipe(new i(n)),e.lock();}catch(e){this._worker=new s("error"),this._worker.error(e);}}f.prototype={accumulate:function(e){return l(this,e)},on:function(e,t){var r=this;return "data"===e?this._worker.on(e,function(e){t.call(r,e.data,e.meta);}):this._worker.on(e,function(){h.delay(t,arguments,r);}),this},resume:function(){return h.delay(this._worker.resume,[],this._worker),this},pause:function(){return this._worker.pause(),this},toNodejsStream:function(e){if(h.checkSupport("nodestream"),"nodebuffer"!==this._outputType)throw new Error(this._outputType+" is not supported by this method");return new o(this,{objectMode:"nodebuffer"!==this._outputType},e)}},t.exports=f;},{"../base64":1,"../external":6,"../nodejs/NodejsStreamOutputAdapter":13,"../support":30,"../utils":32,"./ConvertWorker":24,"./GenericWorker":28}],30:[function(e,t,r){if(r.base64=true,r.array=true,r.string=true,r.arraybuffer="undefined"!=typeof ArrayBuffer&&"undefined"!=typeof Uint8Array,r.nodebuffer="undefined"!=typeof Buffer,r.uint8array="undefined"!=typeof Uint8Array,"undefined"==typeof ArrayBuffer)r.blob=false;else {var n=new ArrayBuffer(0);try{r.blob=0===new Blob([n],{type:"application/zip"}).size;}catch(e){try{var i=new(self.BlobBuilder||self.WebKitBlobBuilder||self.MozBlobBuilder||self.MSBlobBuilder);i.append(n),r.blob=0===i.getBlob("application/zip").size;}catch(e){r.blob=false;}}}try{r.nodestream=!!e("readable-stream").Readable;}catch(e){r.nodestream=false;}},{"readable-stream":16}],31:[function(e,t,s){for(var o=e("./utils"),h=e("./support"),r=e("./nodejsUtils"),n=e("./stream/GenericWorker"),u=new Array(256),i=0;i<256;i++)u[i]=252<=i?6:248<=i?5:240<=i?4:224<=i?3:192<=i?2:1;u[254]=u[254]=1;function a(){n.call(this,"utf-8 decode"),this.leftOver=null;}function l(){n.call(this,"utf-8 encode");}s.utf8encode=function(e){return h.nodebuffer?r.newBufferFrom(e,"utf-8"):function(e){var t,r,n,i,s,a=e.length,o=0;for(i=0;i<a;i++)55296==(64512&(r=e.charCodeAt(i)))&&i+1<a&&56320==(64512&(n=e.charCodeAt(i+1)))&&(r=65536+(r-55296<<10)+(n-56320),i++),o+=r<128?1:r<2048?2:r<65536?3:4;for(t=h.uint8array?new Uint8Array(o):new Array(o),i=s=0;s<o;i++)55296==(64512&(r=e.charCodeAt(i)))&&i+1<a&&56320==(64512&(n=e.charCodeAt(i+1)))&&(r=65536+(r-55296<<10)+(n-56320),i++),r<128?t[s++]=r:(r<2048?t[s++]=192|r>>>6:(r<65536?t[s++]=224|r>>>12:(t[s++]=240|r>>>18,t[s++]=128|r>>>12&63),t[s++]=128|r>>>6&63),t[s++]=128|63&r);return t}(e)},s.utf8decode=function(e){return h.nodebuffer?o.transformTo("nodebuffer",e).toString("utf-8"):function(e){var t,r,n,i,s=e.length,a=new Array(2*s);for(t=r=0;t<s;)if((n=e[t++])<128)a[r++]=n;else if(4<(i=u[n]))a[r++]=65533,t+=i-1;else {for(n&=2===i?31:3===i?15:7;1<i&&t<s;)n=n<<6|63&e[t++],i--;1<i?a[r++]=65533:n<65536?a[r++]=n:(n-=65536,a[r++]=55296|n>>10&1023,a[r++]=56320|1023&n);}return a.length!==r&&(a.subarray?a=a.subarray(0,r):a.length=r),o.applyFromCharCode(a)}(e=o.transformTo(h.uint8array?"uint8array":"array",e))},o.inherits(a,n),a.prototype.processChunk=function(e){var t=o.transformTo(h.uint8array?"uint8array":"array",e.data);if(this.leftOver&&this.leftOver.length){if(h.uint8array){var r=t;(t=new Uint8Array(r.length+this.leftOver.length)).set(this.leftOver,0),t.set(r,this.leftOver.length);}else t=this.leftOver.concat(t);this.leftOver=null;}var n=function(e,t){var r;for((t=t||e.length)>e.length&&(t=e.length),r=t-1;0<=r&&128==(192&e[r]);)r--;return r<0?t:0===r?t:r+u[e[r]]>t?r:t}(t),i=t;n!==t.length&&(h.uint8array?(i=t.subarray(0,n),this.leftOver=t.subarray(n,t.length)):(i=t.slice(0,n),this.leftOver=t.slice(n,t.length))),this.push({data:s.utf8decode(i),meta:e.meta});},a.prototype.flush=function(){this.leftOver&&this.leftOver.length&&(this.push({data:s.utf8decode(this.leftOver),meta:{}}),this.leftOver=null);},s.Utf8DecodeWorker=a,o.inherits(l,n),l.prototype.processChunk=function(e){this.push({data:s.utf8encode(e.data),meta:e.meta});},s.Utf8EncodeWorker=l;},{"./nodejsUtils":14,"./stream/GenericWorker":28,"./support":30,"./utils":32}],32:[function(e,t,a){var o=e("./support"),h=e("./base64"),r=e("./nodejsUtils"),u=e("./external");function n(e){return e}function l(e,t){for(var r=0;r<e.length;++r)t[r]=255&e.charCodeAt(r);return t}e("setimmediate"),a.newBlob=function(t,r){a.checkSupport("blob");try{return new Blob([t],{type:r})}catch(e){try{var n=new(self.BlobBuilder||self.WebKitBlobBuilder||self.MozBlobBuilder||self.MSBlobBuilder);return n.append(t),n.getBlob(r)}catch(e){throw new Error("Bug : can't construct the Blob.")}}};var i={stringifyByChunk:function(e,t,r){var n=[],i=0,s=e.length;if(s<=r)return String.fromCharCode.apply(null,e);for(;i<s;)"array"===t||"nodebuffer"===t?n.push(String.fromCharCode.apply(null,e.slice(i,Math.min(i+r,s)))):n.push(String.fromCharCode.apply(null,e.subarray(i,Math.min(i+r,s)))),i+=r;return n.join("")},stringifyByChar:function(e){for(var t="",r=0;r<e.length;r++)t+=String.fromCharCode(e[r]);return t},applyCanBeUsed:{uint8array:function(){try{return o.uint8array&&1===String.fromCharCode.apply(null,new Uint8Array(1)).length}catch(e){return  false}}(),nodebuffer:function(){try{return o.nodebuffer&&1===String.fromCharCode.apply(null,r.allocBuffer(1)).length}catch(e){return  false}}()}};function s(e){var t=65536,r=a.getTypeOf(e),n=true;if("uint8array"===r?n=i.applyCanBeUsed.uint8array:"nodebuffer"===r&&(n=i.applyCanBeUsed.nodebuffer),n)for(;1<t;)try{return i.stringifyByChunk(e,r,t)}catch(e){t=Math.floor(t/2);}return i.stringifyByChar(e)}function f(e,t){for(var r=0;r<e.length;r++)t[r]=e[r];return t}a.applyFromCharCode=s;var c={};c.string={string:n,array:function(e){return l(e,new Array(e.length))},arraybuffer:function(e){return c.string.uint8array(e).buffer},uint8array:function(e){return l(e,new Uint8Array(e.length))},nodebuffer:function(e){return l(e,r.allocBuffer(e.length))}},c.array={string:s,array:n,arraybuffer:function(e){return new Uint8Array(e).buffer},uint8array:function(e){return new Uint8Array(e)},nodebuffer:function(e){return r.newBufferFrom(e)}},c.arraybuffer={string:function(e){return s(new Uint8Array(e))},array:function(e){return f(new Uint8Array(e),new Array(e.byteLength))},arraybuffer:n,uint8array:function(e){return new Uint8Array(e)},nodebuffer:function(e){return r.newBufferFrom(new Uint8Array(e))}},c.uint8array={string:s,array:function(e){return f(e,new Array(e.length))},arraybuffer:function(e){return e.buffer},uint8array:n,nodebuffer:function(e){return r.newBufferFrom(e)}},c.nodebuffer={string:s,array:function(e){return f(e,new Array(e.length))},arraybuffer:function(e){return c.nodebuffer.uint8array(e).buffer},uint8array:function(e){return f(e,new Uint8Array(e.length))},nodebuffer:n},a.transformTo=function(e,t){if(t=t||"",!e)return t;a.checkSupport(e);var r=a.getTypeOf(t);return c[r][e](t)},a.resolve=function(e){for(var t=e.split("/"),r=[],n=0;n<t.length;n++){var i=t[n];"."===i||""===i&&0!==n&&n!==t.length-1||(".."===i?r.pop():r.push(i));}return r.join("/")},a.getTypeOf=function(e){if("string"==typeof e)return "string";var t=Object.prototype.toString.call(e);return "[object Array]"===t?"array":o.nodebuffer&&r.isBuffer(e)?"nodebuffer":o.uint8array&&"[object Uint8Array]"===t?"uint8array":o.arraybuffer&&"[object ArrayBuffer]"===t?"arraybuffer":void 0},a.checkSupport=function(e){if(!o[e.toLowerCase()])throw new Error(e+" is not supported by this platform")},a.MAX_VALUE_16BITS=65535,a.MAX_VALUE_32BITS=-1,a.pretty=function(e){var t,r,n="";for(r=0;r<(e||"").length;r++)n+="\\x"+((t=e.charCodeAt(r))<16?"0":"")+t.toString(16).toUpperCase();return n},a.delay=function(e,t,r){setImmediate(function(){e.apply(r||null,t||[]);});},a.inherits=function(e,t){function r(){}r.prototype=t.prototype,e.prototype=new r;},a.extend=function(){var e,t,r={};for(e=0;e<arguments.length;e++)for(t in arguments[e])Object.prototype.hasOwnProperty.call(arguments[e],t)&&void 0===r[t]&&(r[t]=arguments[e][t]);return r},a.prepareContent=function(r,e,n,i,s){return u.Promise.resolve(e).then(function(n){return o.blob&&(n instanceof Blob||-1!==["[object File]","[object Blob]"].indexOf(Object.prototype.toString.call(n)))?void 0!==Blob.prototype.arrayBuffer?n.arrayBuffer():"undefined"!=typeof FileReader?new u.Promise(function(t,r){var e=new FileReader;e.onload=function(e){t(e.target.result);},e.onerror=function(e){r(e.target.error);},e.readAsArrayBuffer(n);}):u.Promise.reject(new Error(r+" is a Blob, but we have no way of reading it.")):n}).then(function(e){var t=a.getTypeOf(e);return t?("arraybuffer"===t?e=a.transformTo("uint8array",e):"string"===t&&(s?e=h.decode(e):n&&true!==i&&(e=function(e){return l(e,o.uint8array?new Uint8Array(e.length):new Array(e.length))}(e))),e):u.Promise.reject(new Error("Can't read the data of '"+r+"'. Is it in a supported JavaScript type (String, Blob, ArrayBuffer, etc) ?"))})};},{"./base64":1,"./external":6,"./nodejsUtils":14,"./support":30,setimmediate:54}],33:[function(e,t,r){var n=e("./reader/readerFor"),i=e("./utils"),s=e("./signature"),a=e("./zipEntry"),o=e("./support");function h(e){this.files=[],this.loadOptions=e;}h.prototype={checkSignature:function(e){if(!this.reader.readAndCheckSignature(e)){this.reader.index-=4;var t=this.reader.readString(4);throw new Error("Corrupted zip or bug: unexpected signature ("+i.pretty(t)+", expected "+i.pretty(e)+")")}},isSignature:function(e,t){var r=this.reader.index;this.reader.setIndex(e);var n=this.reader.readString(4)===t;return this.reader.setIndex(r),n},readBlockEndOfCentral:function(){this.diskNumber=this.reader.readInt(2),this.diskWithCentralDirStart=this.reader.readInt(2),this.centralDirRecordsOnThisDisk=this.reader.readInt(2),this.centralDirRecords=this.reader.readInt(2),this.centralDirSize=this.reader.readInt(4),this.centralDirOffset=this.reader.readInt(4),this.zipCommentLength=this.reader.readInt(2);var e=this.reader.readData(this.zipCommentLength),t=o.uint8array?"uint8array":"array",r=i.transformTo(t,e);this.zipComment=this.loadOptions.decodeFileName(r);},readBlockZip64EndOfCentral:function(){this.zip64EndOfCentralSize=this.reader.readInt(8),this.reader.skip(4),this.diskNumber=this.reader.readInt(4),this.diskWithCentralDirStart=this.reader.readInt(4),this.centralDirRecordsOnThisDisk=this.reader.readInt(8),this.centralDirRecords=this.reader.readInt(8),this.centralDirSize=this.reader.readInt(8),this.centralDirOffset=this.reader.readInt(8),this.zip64ExtensibleData={};for(var e,t,r,n=this.zip64EndOfCentralSize-44;0<n;)e=this.reader.readInt(2),t=this.reader.readInt(4),r=this.reader.readData(t),this.zip64ExtensibleData[e]={id:e,length:t,value:r};},readBlockZip64EndOfCentralLocator:function(){if(this.diskWithZip64CentralDirStart=this.reader.readInt(4),this.relativeOffsetEndOfZip64CentralDir=this.reader.readInt(8),this.disksCount=this.reader.readInt(4),1<this.disksCount)throw new Error("Multi-volumes zip are not supported")},readLocalFiles:function(){var e,t;for(e=0;e<this.files.length;e++)t=this.files[e],this.reader.setIndex(t.localHeaderOffset),this.checkSignature(s.LOCAL_FILE_HEADER),t.readLocalPart(this.reader),t.handleUTF8(),t.processAttributes();},readCentralDir:function(){var e;for(this.reader.setIndex(this.centralDirOffset);this.reader.readAndCheckSignature(s.CENTRAL_FILE_HEADER);)(e=new a({zip64:this.zip64},this.loadOptions)).readCentralPart(this.reader),this.files.push(e);if(this.centralDirRecords!==this.files.length&&0!==this.centralDirRecords&&0===this.files.length)throw new Error("Corrupted zip or bug: expected "+this.centralDirRecords+" records in central dir, got "+this.files.length)},readEndOfCentral:function(){var e=this.reader.lastIndexOfSignature(s.CENTRAL_DIRECTORY_END);if(e<0)throw !this.isSignature(0,s.LOCAL_FILE_HEADER)?new Error("Can't find end of central directory : is this a zip file ? If it is, see https://stuk.github.io/jszip/documentation/howto/read_zip.html"):new Error("Corrupted zip: can't find end of central directory");this.reader.setIndex(e);var t=e;if(this.checkSignature(s.CENTRAL_DIRECTORY_END),this.readBlockEndOfCentral(),this.diskNumber===i.MAX_VALUE_16BITS||this.diskWithCentralDirStart===i.MAX_VALUE_16BITS||this.centralDirRecordsOnThisDisk===i.MAX_VALUE_16BITS||this.centralDirRecords===i.MAX_VALUE_16BITS||this.centralDirSize===i.MAX_VALUE_32BITS||this.centralDirOffset===i.MAX_VALUE_32BITS){if(this.zip64=true,(e=this.reader.lastIndexOfSignature(s.ZIP64_CENTRAL_DIRECTORY_LOCATOR))<0)throw new Error("Corrupted zip: can't find the ZIP64 end of central directory locator");if(this.reader.setIndex(e),this.checkSignature(s.ZIP64_CENTRAL_DIRECTORY_LOCATOR),this.readBlockZip64EndOfCentralLocator(),!this.isSignature(this.relativeOffsetEndOfZip64CentralDir,s.ZIP64_CENTRAL_DIRECTORY_END)&&(this.relativeOffsetEndOfZip64CentralDir=this.reader.lastIndexOfSignature(s.ZIP64_CENTRAL_DIRECTORY_END),this.relativeOffsetEndOfZip64CentralDir<0))throw new Error("Corrupted zip: can't find the ZIP64 end of central directory");this.reader.setIndex(this.relativeOffsetEndOfZip64CentralDir),this.checkSignature(s.ZIP64_CENTRAL_DIRECTORY_END),this.readBlockZip64EndOfCentral();}var r=this.centralDirOffset+this.centralDirSize;this.zip64&&(r+=20,r+=12+this.zip64EndOfCentralSize);var n=t-r;if(0<n)this.isSignature(t,s.CENTRAL_FILE_HEADER)||(this.reader.zero=n);else if(n<0)throw new Error("Corrupted zip: missing "+Math.abs(n)+" bytes.")},prepareReader:function(e){this.reader=n(e);},load:function(e){this.prepareReader(e),this.readEndOfCentral(),this.readCentralDir(),this.readLocalFiles();}},t.exports=h;},{"./reader/readerFor":22,"./signature":23,"./support":30,"./utils":32,"./zipEntry":34}],34:[function(e,t,r){var n=e("./reader/readerFor"),s=e("./utils"),i=e("./compressedObject"),a=e("./crc32"),o=e("./utf8"),h=e("./compressions"),u=e("./support");function l(e,t){this.options=e,this.loadOptions=t;}l.prototype={isEncrypted:function(){return 1==(1&this.bitFlag)},useUTF8:function(){return 2048==(2048&this.bitFlag)},readLocalPart:function(e){var t,r;if(e.skip(22),this.fileNameLength=e.readInt(2),r=e.readInt(2),this.fileName=e.readData(this.fileNameLength),e.skip(r),-1===this.compressedSize||-1===this.uncompressedSize)throw new Error("Bug or corrupted zip : didn't get enough information from the central directory (compressedSize === -1 || uncompressedSize === -1)");if(null===(t=function(e){for(var t in h)if(Object.prototype.hasOwnProperty.call(h,t)&&h[t].magic===e)return h[t];return null}(this.compressionMethod)))throw new Error("Corrupted zip : compression "+s.pretty(this.compressionMethod)+" unknown (inner file : "+s.transformTo("string",this.fileName)+")");this.decompressed=new i(this.compressedSize,this.uncompressedSize,this.crc32,t,e.readData(this.compressedSize));},readCentralPart:function(e){this.versionMadeBy=e.readInt(2),e.skip(2),this.bitFlag=e.readInt(2),this.compressionMethod=e.readString(2),this.date=e.readDate(),this.crc32=e.readInt(4),this.compressedSize=e.readInt(4),this.uncompressedSize=e.readInt(4);var t=e.readInt(2);if(this.extraFieldsLength=e.readInt(2),this.fileCommentLength=e.readInt(2),this.diskNumberStart=e.readInt(2),this.internalFileAttributes=e.readInt(2),this.externalFileAttributes=e.readInt(4),this.localHeaderOffset=e.readInt(4),this.isEncrypted())throw new Error("Encrypted zip are not supported");e.skip(t),this.readExtraFields(e),this.parseZIP64ExtraField(e),this.fileComment=e.readData(this.fileCommentLength);},processAttributes:function(){this.unixPermissions=null,this.dosPermissions=null;var e=this.versionMadeBy>>8;this.dir=!!(16&this.externalFileAttributes),0==e&&(this.dosPermissions=63&this.externalFileAttributes),3==e&&(this.unixPermissions=this.externalFileAttributes>>16&65535),this.dir||"/"!==this.fileNameStr.slice(-1)||(this.dir=true);},parseZIP64ExtraField:function(){if(this.extraFields[1]){var e=n(this.extraFields[1].value);this.uncompressedSize===s.MAX_VALUE_32BITS&&(this.uncompressedSize=e.readInt(8)),this.compressedSize===s.MAX_VALUE_32BITS&&(this.compressedSize=e.readInt(8)),this.localHeaderOffset===s.MAX_VALUE_32BITS&&(this.localHeaderOffset=e.readInt(8)),this.diskNumberStart===s.MAX_VALUE_32BITS&&(this.diskNumberStart=e.readInt(4));}},readExtraFields:function(e){var t,r,n,i=e.index+this.extraFieldsLength;for(this.extraFields||(this.extraFields={});e.index+4<i;)t=e.readInt(2),r=e.readInt(2),n=e.readData(r),this.extraFields[t]={id:t,length:r,value:n};e.setIndex(i);},handleUTF8:function(){var e=u.uint8array?"uint8array":"array";if(this.useUTF8())this.fileNameStr=o.utf8decode(this.fileName),this.fileCommentStr=o.utf8decode(this.fileComment);else {var t=this.findExtraFieldUnicodePath();if(null!==t)this.fileNameStr=t;else {var r=s.transformTo(e,this.fileName);this.fileNameStr=this.loadOptions.decodeFileName(r);}var n=this.findExtraFieldUnicodeComment();if(null!==n)this.fileCommentStr=n;else {var i=s.transformTo(e,this.fileComment);this.fileCommentStr=this.loadOptions.decodeFileName(i);}}},findExtraFieldUnicodePath:function(){var e=this.extraFields[28789];if(e){var t=n(e.value);return 1!==t.readInt(1)?null:a(this.fileName)!==t.readInt(4)?null:o.utf8decode(t.readData(e.length-5))}return null},findExtraFieldUnicodeComment:function(){var e=this.extraFields[25461];if(e){var t=n(e.value);return 1!==t.readInt(1)?null:a(this.fileComment)!==t.readInt(4)?null:o.utf8decode(t.readData(e.length-5))}return null}},t.exports=l;},{"./compressedObject":2,"./compressions":3,"./crc32":4,"./reader/readerFor":22,"./support":30,"./utf8":31,"./utils":32}],35:[function(e,t,r){function n(e,t,r){this.name=e,this.dir=r.dir,this.date=r.date,this.comment=r.comment,this.unixPermissions=r.unixPermissions,this.dosPermissions=r.dosPermissions,this._data=t,this._dataBinary=r.binary,this.options={compression:r.compression,compressionOptions:r.compressionOptions};}var s=e("./stream/StreamHelper"),i=e("./stream/DataWorker"),a=e("./utf8"),o=e("./compressedObject"),h=e("./stream/GenericWorker");n.prototype={internalStream:function(e){var t=null,r="string";try{if(!e)throw new Error("No output type specified.");var n="string"===(r=e.toLowerCase())||"text"===r;"binarystring"!==r&&"text"!==r||(r="string"),t=this._decompressWorker();var i=!this._dataBinary;i&&!n&&(t=t.pipe(new a.Utf8EncodeWorker)),!i&&n&&(t=t.pipe(new a.Utf8DecodeWorker));}catch(e){(t=new h("error")).error(e);}return new s(t,r,"")},async:function(e,t){return this.internalStream(e).accumulate(t)},nodeStream:function(e,t){return this.internalStream(e||"nodebuffer").toNodejsStream(t)},_compressWorker:function(e,t){if(this._data instanceof o&&this._data.compression.magic===e.magic)return this._data.getCompressedWorker();var r=this._decompressWorker();return this._dataBinary||(r=r.pipe(new a.Utf8EncodeWorker)),o.createWorkerFrom(r,e,t)},_decompressWorker:function(){return this._data instanceof o?this._data.getContentWorker():this._data instanceof h?this._data:new i(this._data)}};for(var u=["asText","asBinary","asNodeBuffer","asUint8Array","asArrayBuffer"],l=function(){throw new Error("This method has been removed in JSZip 3.0, please check the upgrade guide.")},f=0;f<u.length;f++)n.prototype[u[f]]=l;t.exports=n;},{"./compressedObject":2,"./stream/DataWorker":27,"./stream/GenericWorker":28,"./stream/StreamHelper":29,"./utf8":31}],36:[function(e,l,t){(function(t){var r,n,e=t.MutationObserver||t.WebKitMutationObserver;if(e){var i=0,s=new e(u),a=t.document.createTextNode("");s.observe(a,{characterData:true}),r=function(){a.data=i=++i%2;};}else if(t.setImmediate||void 0===t.MessageChannel)r="document"in t&&"onreadystatechange"in t.document.createElement("script")?function(){var e=t.document.createElement("script");e.onreadystatechange=function(){u(),e.onreadystatechange=null,e.parentNode.removeChild(e),e=null;},t.document.documentElement.appendChild(e);}:function(){setTimeout(u,0);};else {var o=new t.MessageChannel;o.port1.onmessage=u,r=function(){o.port2.postMessage(0);};}var h=[];function u(){var e,t;n=true;for(var r=h.length;r;){for(t=h,h=[],e=-1;++e<r;)t[e]();r=h.length;}n=false;}l.exports=function(e){1!==h.push(e)||n||r();};}).call(this,"undefined"!=typeof commonjsGlobal?commonjsGlobal:"undefined"!=typeof self?self:"undefined"!=typeof window?window:{});},{}],37:[function(e,t,r){var i=e("immediate");function u(){}var l={},s=["REJECTED"],a=["FULFILLED"],n=["PENDING"];function o(e){if("function"!=typeof e)throw new TypeError("resolver must be a function");this.state=n,this.queue=[],this.outcome=void 0,e!==u&&d(this,e);}function h(e,t,r){this.promise=e,"function"==typeof t&&(this.onFulfilled=t,this.callFulfilled=this.otherCallFulfilled),"function"==typeof r&&(this.onRejected=r,this.callRejected=this.otherCallRejected);}function f(t,r,n){i(function(){var e;try{e=r(n);}catch(e){return l.reject(t,e)}e===t?l.reject(t,new TypeError("Cannot resolve promise with itself")):l.resolve(t,e);});}function c(e){var t=e&&e.then;if(e&&("object"==typeof e||"function"==typeof e)&&"function"==typeof t)return function(){t.apply(e,arguments);}}function d(t,e){var r=false;function n(e){r||(r=true,l.reject(t,e));}function i(e){r||(r=true,l.resolve(t,e));}var s=p(function(){e(i,n);});"error"===s.status&&n(s.value);}function p(e,t){var r={};try{r.value=e(t),r.status="success";}catch(e){r.status="error",r.value=e;}return r}(t.exports=o).prototype.finally=function(t){if("function"!=typeof t)return this;var r=this.constructor;return this.then(function(e){return r.resolve(t()).then(function(){return e})},function(e){return r.resolve(t()).then(function(){throw e})})},o.prototype.catch=function(e){return this.then(null,e)},o.prototype.then=function(e,t){if("function"!=typeof e&&this.state===a||"function"!=typeof t&&this.state===s)return this;var r=new this.constructor(u);this.state!==n?f(r,this.state===a?e:t,this.outcome):this.queue.push(new h(r,e,t));return r},h.prototype.callFulfilled=function(e){l.resolve(this.promise,e);},h.prototype.otherCallFulfilled=function(e){f(this.promise,this.onFulfilled,e);},h.prototype.callRejected=function(e){l.reject(this.promise,e);},h.prototype.otherCallRejected=function(e){f(this.promise,this.onRejected,e);},l.resolve=function(e,t){var r=p(c,t);if("error"===r.status)return l.reject(e,r.value);var n=r.value;if(n)d(e,n);else {e.state=a,e.outcome=t;for(var i=-1,s=e.queue.length;++i<s;)e.queue[i].callFulfilled(t);}return e},l.reject=function(e,t){e.state=s,e.outcome=t;for(var r=-1,n=e.queue.length;++r<n;)e.queue[r].callRejected(t);return e},o.resolve=function(e){if(e instanceof this)return e;return l.resolve(new this(u),e)},o.reject=function(e){var t=new this(u);return l.reject(t,e)},o.all=function(e){var r=this;if("[object Array]"!==Object.prototype.toString.call(e))return this.reject(new TypeError("must be an array"));var n=e.length,i=false;if(!n)return this.resolve([]);var s=new Array(n),a=0,t=-1,o=new this(u);for(;++t<n;)h(e[t],t);return o;function h(e,t){r.resolve(e).then(function(e){s[t]=e,++a!==n||i||(i=true,l.resolve(o,s));},function(e){i||(i=true,l.reject(o,e));});}},o.race=function(e){var t=this;if("[object Array]"!==Object.prototype.toString.call(e))return this.reject(new TypeError("must be an array"));var r=e.length,n=false;if(!r)return this.resolve([]);var i=-1,s=new this(u);for(;++i<r;)a=e[i],t.resolve(a).then(function(e){n||(n=true,l.resolve(s,e));},function(e){n||(n=true,l.reject(s,e));});var a;return s};},{immediate:36}],38:[function(e,t,r){var n={};(0, e("./lib/utils/common").assign)(n,e("./lib/deflate"),e("./lib/inflate"),e("./lib/zlib/constants")),t.exports=n;},{"./lib/deflate":39,"./lib/inflate":40,"./lib/utils/common":41,"./lib/zlib/constants":44}],39:[function(e,t,r){var a=e("./zlib/deflate"),o=e("./utils/common"),h=e("./utils/strings"),i=e("./zlib/messages"),s=e("./zlib/zstream"),u=Object.prototype.toString,l=0,f=-1,c=0,d=8;function p(e){if(!(this instanceof p))return new p(e);this.options=o.assign({level:f,method:d,chunkSize:16384,windowBits:15,memLevel:8,strategy:c,to:""},e||{});var t=this.options;t.raw&&0<t.windowBits?t.windowBits=-t.windowBits:t.gzip&&0<t.windowBits&&t.windowBits<16&&(t.windowBits+=16),this.err=0,this.msg="",this.ended=false,this.chunks=[],this.strm=new s,this.strm.avail_out=0;var r=a.deflateInit2(this.strm,t.level,t.method,t.windowBits,t.memLevel,t.strategy);if(r!==l)throw new Error(i[r]);if(t.header&&a.deflateSetHeader(this.strm,t.header),t.dictionary){var n;if(n="string"==typeof t.dictionary?h.string2buf(t.dictionary):"[object ArrayBuffer]"===u.call(t.dictionary)?new Uint8Array(t.dictionary):t.dictionary,(r=a.deflateSetDictionary(this.strm,n))!==l)throw new Error(i[r]);this._dict_set=true;}}function n(e,t){var r=new p(t);if(r.push(e,true),r.err)throw r.msg||i[r.err];return r.result}p.prototype.push=function(e,t){var r,n,i=this.strm,s=this.options.chunkSize;if(this.ended)return  false;n=t===~~t?t:true===t?4:0,"string"==typeof e?i.input=h.string2buf(e):"[object ArrayBuffer]"===u.call(e)?i.input=new Uint8Array(e):i.input=e,i.next_in=0,i.avail_in=i.input.length;do{if(0===i.avail_out&&(i.output=new o.Buf8(s),i.next_out=0,i.avail_out=s),1!==(r=a.deflate(i,n))&&r!==l)return this.onEnd(r),!(this.ended=true);0!==i.avail_out&&(0!==i.avail_in||4!==n&&2!==n)||("string"===this.options.to?this.onData(h.buf2binstring(o.shrinkBuf(i.output,i.next_out))):this.onData(o.shrinkBuf(i.output,i.next_out)));}while((0<i.avail_in||0===i.avail_out)&&1!==r);return 4===n?(r=a.deflateEnd(this.strm),this.onEnd(r),this.ended=true,r===l):2!==n||(this.onEnd(l),!(i.avail_out=0))},p.prototype.onData=function(e){this.chunks.push(e);},p.prototype.onEnd=function(e){e===l&&("string"===this.options.to?this.result=this.chunks.join(""):this.result=o.flattenChunks(this.chunks)),this.chunks=[],this.err=e,this.msg=this.strm.msg;},r.Deflate=p,r.deflate=n,r.deflateRaw=function(e,t){return (t=t||{}).raw=true,n(e,t)},r.gzip=function(e,t){return (t=t||{}).gzip=true,n(e,t)};},{"./utils/common":41,"./utils/strings":42,"./zlib/deflate":46,"./zlib/messages":51,"./zlib/zstream":53}],40:[function(e,t,r){var c=e("./zlib/inflate"),d=e("./utils/common"),p=e("./utils/strings"),m=e("./zlib/constants"),n=e("./zlib/messages"),i=e("./zlib/zstream"),s=e("./zlib/gzheader"),_=Object.prototype.toString;function a(e){if(!(this instanceof a))return new a(e);this.options=d.assign({chunkSize:16384,windowBits:0,to:""},e||{});var t=this.options;t.raw&&0<=t.windowBits&&t.windowBits<16&&(t.windowBits=-t.windowBits,0===t.windowBits&&(t.windowBits=-15)),!(0<=t.windowBits&&t.windowBits<16)||e&&e.windowBits||(t.windowBits+=32),15<t.windowBits&&t.windowBits<48&&0==(15&t.windowBits)&&(t.windowBits|=15),this.err=0,this.msg="",this.ended=false,this.chunks=[],this.strm=new i,this.strm.avail_out=0;var r=c.inflateInit2(this.strm,t.windowBits);if(r!==m.Z_OK)throw new Error(n[r]);this.header=new s,c.inflateGetHeader(this.strm,this.header);}function o(e,t){var r=new a(t);if(r.push(e,true),r.err)throw r.msg||n[r.err];return r.result}a.prototype.push=function(e,t){var r,n,i,s,a,o,h=this.strm,u=this.options.chunkSize,l=this.options.dictionary,f=false;if(this.ended)return  false;n=t===~~t?t:true===t?m.Z_FINISH:m.Z_NO_FLUSH,"string"==typeof e?h.input=p.binstring2buf(e):"[object ArrayBuffer]"===_.call(e)?h.input=new Uint8Array(e):h.input=e,h.next_in=0,h.avail_in=h.input.length;do{if(0===h.avail_out&&(h.output=new d.Buf8(u),h.next_out=0,h.avail_out=u),(r=c.inflate(h,m.Z_NO_FLUSH))===m.Z_NEED_DICT&&l&&(o="string"==typeof l?p.string2buf(l):"[object ArrayBuffer]"===_.call(l)?new Uint8Array(l):l,r=c.inflateSetDictionary(this.strm,o)),r===m.Z_BUF_ERROR&&true===f&&(r=m.Z_OK,f=false),r!==m.Z_STREAM_END&&r!==m.Z_OK)return this.onEnd(r),!(this.ended=true);h.next_out&&(0!==h.avail_out&&r!==m.Z_STREAM_END&&(0!==h.avail_in||n!==m.Z_FINISH&&n!==m.Z_SYNC_FLUSH)||("string"===this.options.to?(i=p.utf8border(h.output,h.next_out),s=h.next_out-i,a=p.buf2string(h.output,i),h.next_out=s,h.avail_out=u-s,s&&d.arraySet(h.output,h.output,i,s,0),this.onData(a)):this.onData(d.shrinkBuf(h.output,h.next_out)))),0===h.avail_in&&0===h.avail_out&&(f=true);}while((0<h.avail_in||0===h.avail_out)&&r!==m.Z_STREAM_END);return r===m.Z_STREAM_END&&(n=m.Z_FINISH),n===m.Z_FINISH?(r=c.inflateEnd(this.strm),this.onEnd(r),this.ended=true,r===m.Z_OK):n!==m.Z_SYNC_FLUSH||(this.onEnd(m.Z_OK),!(h.avail_out=0))},a.prototype.onData=function(e){this.chunks.push(e);},a.prototype.onEnd=function(e){e===m.Z_OK&&("string"===this.options.to?this.result=this.chunks.join(""):this.result=d.flattenChunks(this.chunks)),this.chunks=[],this.err=e,this.msg=this.strm.msg;},r.Inflate=a,r.inflate=o,r.inflateRaw=function(e,t){return (t=t||{}).raw=true,o(e,t)},r.ungzip=o;},{"./utils/common":41,"./utils/strings":42,"./zlib/constants":44,"./zlib/gzheader":47,"./zlib/inflate":49,"./zlib/messages":51,"./zlib/zstream":53}],41:[function(e,t,r){var n="undefined"!=typeof Uint8Array&&"undefined"!=typeof Uint16Array&&"undefined"!=typeof Int32Array;r.assign=function(e){for(var t=Array.prototype.slice.call(arguments,1);t.length;){var r=t.shift();if(r){if("object"!=typeof r)throw new TypeError(r+"must be non-object");for(var n in r)r.hasOwnProperty(n)&&(e[n]=r[n]);}}return e},r.shrinkBuf=function(e,t){return e.length===t?e:e.subarray?e.subarray(0,t):(e.length=t,e)};var i={arraySet:function(e,t,r,n,i){if(t.subarray&&e.subarray)e.set(t.subarray(r,r+n),i);else for(var s=0;s<n;s++)e[i+s]=t[r+s];},flattenChunks:function(e){var t,r,n,i,s,a;for(t=n=0,r=e.length;t<r;t++)n+=e[t].length;for(a=new Uint8Array(n),t=i=0,r=e.length;t<r;t++)s=e[t],a.set(s,i),i+=s.length;return a}},s={arraySet:function(e,t,r,n,i){for(var s=0;s<n;s++)e[i+s]=t[r+s];},flattenChunks:function(e){return [].concat.apply([],e)}};r.setTyped=function(e){e?(r.Buf8=Uint8Array,r.Buf16=Uint16Array,r.Buf32=Int32Array,r.assign(r,i)):(r.Buf8=Array,r.Buf16=Array,r.Buf32=Array,r.assign(r,s));},r.setTyped(n);},{}],42:[function(e,t,r){var h=e("./common"),i=true,s=true;try{String.fromCharCode.apply(null,[0]);}catch(e){i=false;}try{String.fromCharCode.apply(null,new Uint8Array(1));}catch(e){s=false;}for(var u=new h.Buf8(256),n=0;n<256;n++)u[n]=252<=n?6:248<=n?5:240<=n?4:224<=n?3:192<=n?2:1;function l(e,t){if(t<65537&&(e.subarray&&s||!e.subarray&&i))return String.fromCharCode.apply(null,h.shrinkBuf(e,t));for(var r="",n=0;n<t;n++)r+=String.fromCharCode(e[n]);return r}u[254]=u[254]=1,r.string2buf=function(e){var t,r,n,i,s,a=e.length,o=0;for(i=0;i<a;i++)55296==(64512&(r=e.charCodeAt(i)))&&i+1<a&&56320==(64512&(n=e.charCodeAt(i+1)))&&(r=65536+(r-55296<<10)+(n-56320),i++),o+=r<128?1:r<2048?2:r<65536?3:4;for(t=new h.Buf8(o),i=s=0;s<o;i++)55296==(64512&(r=e.charCodeAt(i)))&&i+1<a&&56320==(64512&(n=e.charCodeAt(i+1)))&&(r=65536+(r-55296<<10)+(n-56320),i++),r<128?t[s++]=r:(r<2048?t[s++]=192|r>>>6:(r<65536?t[s++]=224|r>>>12:(t[s++]=240|r>>>18,t[s++]=128|r>>>12&63),t[s++]=128|r>>>6&63),t[s++]=128|63&r);return t},r.buf2binstring=function(e){return l(e,e.length)},r.binstring2buf=function(e){for(var t=new h.Buf8(e.length),r=0,n=t.length;r<n;r++)t[r]=e.charCodeAt(r);return t},r.buf2string=function(e,t){var r,n,i,s,a=t||e.length,o=new Array(2*a);for(r=n=0;r<a;)if((i=e[r++])<128)o[n++]=i;else if(4<(s=u[i]))o[n++]=65533,r+=s-1;else {for(i&=2===s?31:3===s?15:7;1<s&&r<a;)i=i<<6|63&e[r++],s--;1<s?o[n++]=65533:i<65536?o[n++]=i:(i-=65536,o[n++]=55296|i>>10&1023,o[n++]=56320|1023&i);}return l(o,n)},r.utf8border=function(e,t){var r;for((t=t||e.length)>e.length&&(t=e.length),r=t-1;0<=r&&128==(192&e[r]);)r--;return r<0?t:0===r?t:r+u[e[r]]>t?r:t};},{"./common":41}],43:[function(e,t,r){t.exports=function(e,t,r,n){for(var i=65535&e|0,s=e>>>16&65535|0,a=0;0!==r;){for(r-=a=2e3<r?2e3:r;s=s+(i=i+t[n++]|0)|0,--a;);i%=65521,s%=65521;}return i|s<<16|0};},{}],44:[function(e,t,r){t.exports={Z_NO_FLUSH:0,Z_PARTIAL_FLUSH:1,Z_SYNC_FLUSH:2,Z_FULL_FLUSH:3,Z_FINISH:4,Z_BLOCK:5,Z_TREES:6,Z_OK:0,Z_STREAM_END:1,Z_NEED_DICT:2,Z_ERRNO:-1,Z_STREAM_ERROR:-2,Z_DATA_ERROR:-3,Z_BUF_ERROR:-5,Z_NO_COMPRESSION:0,Z_BEST_SPEED:1,Z_BEST_COMPRESSION:9,Z_DEFAULT_COMPRESSION:-1,Z_FILTERED:1,Z_HUFFMAN_ONLY:2,Z_RLE:3,Z_FIXED:4,Z_DEFAULT_STRATEGY:0,Z_BINARY:0,Z_TEXT:1,Z_UNKNOWN:2,Z_DEFLATED:8};},{}],45:[function(e,t,r){var o=function(){for(var e,t=[],r=0;r<256;r++){e=r;for(var n=0;n<8;n++)e=1&e?3988292384^e>>>1:e>>>1;t[r]=e;}return t}();t.exports=function(e,t,r,n){var i=o,s=n+r;e^=-1;for(var a=n;a<s;a++)e=e>>>8^i[255&(e^t[a])];return  -1^e};},{}],46:[function(e,t,r){var h,c=e("../utils/common"),u=e("./trees"),d=e("./adler32"),p=e("./crc32"),n=e("./messages"),l=0,f=4,m=0,_=-2,g=-1,b=4,i=2,v=8,y=9,s=286,a=30,o=19,w=2*s+1,k=15,x=3,S=258,z=S+x+1,C=42,E=113,A=1,I=2,O=3,B=4;function R(e,t){return e.msg=n[t],t}function T(e){return (e<<1)-(4<e?9:0)}function D(e){for(var t=e.length;0<=--t;)e[t]=0;}function F(e){var t=e.state,r=t.pending;r>e.avail_out&&(r=e.avail_out),0!==r&&(c.arraySet(e.output,t.pending_buf,t.pending_out,r,e.next_out),e.next_out+=r,t.pending_out+=r,e.total_out+=r,e.avail_out-=r,t.pending-=r,0===t.pending&&(t.pending_out=0));}function N(e,t){u._tr_flush_block(e,0<=e.block_start?e.block_start:-1,e.strstart-e.block_start,t),e.block_start=e.strstart,F(e.strm);}function U(e,t){e.pending_buf[e.pending++]=t;}function P(e,t){e.pending_buf[e.pending++]=t>>>8&255,e.pending_buf[e.pending++]=255&t;}function L(e,t){var r,n,i=e.max_chain_length,s=e.strstart,a=e.prev_length,o=e.nice_match,h=e.strstart>e.w_size-z?e.strstart-(e.w_size-z):0,u=e.window,l=e.w_mask,f=e.prev,c=e.strstart+S,d=u[s+a-1],p=u[s+a];e.prev_length>=e.good_match&&(i>>=2),o>e.lookahead&&(o=e.lookahead);do{if(u[(r=t)+a]===p&&u[r+a-1]===d&&u[r]===u[s]&&u[++r]===u[s+1]){s+=2,r++;do{}while(u[++s]===u[++r]&&u[++s]===u[++r]&&u[++s]===u[++r]&&u[++s]===u[++r]&&u[++s]===u[++r]&&u[++s]===u[++r]&&u[++s]===u[++r]&&u[++s]===u[++r]&&s<c);if(n=S-(c-s),s=c-S,a<n){if(e.match_start=t,o<=(a=n))break;d=u[s+a-1],p=u[s+a];}}}while((t=f[t&l])>h&&0!=--i);return a<=e.lookahead?a:e.lookahead}function j(e){var t,r,n,i,s,a,o,h,u,l,f=e.w_size;do{if(i=e.window_size-e.lookahead-e.strstart,e.strstart>=f+(f-z)){for(c.arraySet(e.window,e.window,f,f,0),e.match_start-=f,e.strstart-=f,e.block_start-=f,t=r=e.hash_size;n=e.head[--t],e.head[t]=f<=n?n-f:0,--r;);for(t=r=f;n=e.prev[--t],e.prev[t]=f<=n?n-f:0,--r;);i+=f;}if(0===e.strm.avail_in)break;if(a=e.strm,o=e.window,h=e.strstart+e.lookahead,u=i,l=void 0,l=a.avail_in,u<l&&(l=u),r=0===l?0:(a.avail_in-=l,c.arraySet(o,a.input,a.next_in,l,h),1===a.state.wrap?a.adler=d(a.adler,o,l,h):2===a.state.wrap&&(a.adler=p(a.adler,o,l,h)),a.next_in+=l,a.total_in+=l,l),e.lookahead+=r,e.lookahead+e.insert>=x)for(s=e.strstart-e.insert,e.ins_h=e.window[s],e.ins_h=(e.ins_h<<e.hash_shift^e.window[s+1])&e.hash_mask;e.insert&&(e.ins_h=(e.ins_h<<e.hash_shift^e.window[s+x-1])&e.hash_mask,e.prev[s&e.w_mask]=e.head[e.ins_h],e.head[e.ins_h]=s,s++,e.insert--,!(e.lookahead+e.insert<x)););}while(e.lookahead<z&&0!==e.strm.avail_in)}function Z(e,t){for(var r,n;;){if(e.lookahead<z){if(j(e),e.lookahead<z&&t===l)return A;if(0===e.lookahead)break}if(r=0,e.lookahead>=x&&(e.ins_h=(e.ins_h<<e.hash_shift^e.window[e.strstart+x-1])&e.hash_mask,r=e.prev[e.strstart&e.w_mask]=e.head[e.ins_h],e.head[e.ins_h]=e.strstart),0!==r&&e.strstart-r<=e.w_size-z&&(e.match_length=L(e,r)),e.match_length>=x)if(n=u._tr_tally(e,e.strstart-e.match_start,e.match_length-x),e.lookahead-=e.match_length,e.match_length<=e.max_lazy_match&&e.lookahead>=x){for(e.match_length--;e.strstart++,e.ins_h=(e.ins_h<<e.hash_shift^e.window[e.strstart+x-1])&e.hash_mask,r=e.prev[e.strstart&e.w_mask]=e.head[e.ins_h],e.head[e.ins_h]=e.strstart,0!=--e.match_length;);e.strstart++;}else e.strstart+=e.match_length,e.match_length=0,e.ins_h=e.window[e.strstart],e.ins_h=(e.ins_h<<e.hash_shift^e.window[e.strstart+1])&e.hash_mask;else n=u._tr_tally(e,0,e.window[e.strstart]),e.lookahead--,e.strstart++;if(n&&(N(e,false),0===e.strm.avail_out))return A}return e.insert=e.strstart<x-1?e.strstart:x-1,t===f?(N(e,true),0===e.strm.avail_out?O:B):e.last_lit&&(N(e,false),0===e.strm.avail_out)?A:I}function W(e,t){for(var r,n,i;;){if(e.lookahead<z){if(j(e),e.lookahead<z&&t===l)return A;if(0===e.lookahead)break}if(r=0,e.lookahead>=x&&(e.ins_h=(e.ins_h<<e.hash_shift^e.window[e.strstart+x-1])&e.hash_mask,r=e.prev[e.strstart&e.w_mask]=e.head[e.ins_h],e.head[e.ins_h]=e.strstart),e.prev_length=e.match_length,e.prev_match=e.match_start,e.match_length=x-1,0!==r&&e.prev_length<e.max_lazy_match&&e.strstart-r<=e.w_size-z&&(e.match_length=L(e,r),e.match_length<=5&&(1===e.strategy||e.match_length===x&&4096<e.strstart-e.match_start)&&(e.match_length=x-1)),e.prev_length>=x&&e.match_length<=e.prev_length){for(i=e.strstart+e.lookahead-x,n=u._tr_tally(e,e.strstart-1-e.prev_match,e.prev_length-x),e.lookahead-=e.prev_length-1,e.prev_length-=2;++e.strstart<=i&&(e.ins_h=(e.ins_h<<e.hash_shift^e.window[e.strstart+x-1])&e.hash_mask,r=e.prev[e.strstart&e.w_mask]=e.head[e.ins_h],e.head[e.ins_h]=e.strstart),0!=--e.prev_length;);if(e.match_available=0,e.match_length=x-1,e.strstart++,n&&(N(e,false),0===e.strm.avail_out))return A}else if(e.match_available){if((n=u._tr_tally(e,0,e.window[e.strstart-1]))&&N(e,false),e.strstart++,e.lookahead--,0===e.strm.avail_out)return A}else e.match_available=1,e.strstart++,e.lookahead--;}return e.match_available&&(n=u._tr_tally(e,0,e.window[e.strstart-1]),e.match_available=0),e.insert=e.strstart<x-1?e.strstart:x-1,t===f?(N(e,true),0===e.strm.avail_out?O:B):e.last_lit&&(N(e,false),0===e.strm.avail_out)?A:I}function M(e,t,r,n,i){this.good_length=e,this.max_lazy=t,this.nice_length=r,this.max_chain=n,this.func=i;}function H(){this.strm=null,this.status=0,this.pending_buf=null,this.pending_buf_size=0,this.pending_out=0,this.pending=0,this.wrap=0,this.gzhead=null,this.gzindex=0,this.method=v,this.last_flush=-1,this.w_size=0,this.w_bits=0,this.w_mask=0,this.window=null,this.window_size=0,this.prev=null,this.head=null,this.ins_h=0,this.hash_size=0,this.hash_bits=0,this.hash_mask=0,this.hash_shift=0,this.block_start=0,this.match_length=0,this.prev_match=0,this.match_available=0,this.strstart=0,this.match_start=0,this.lookahead=0,this.prev_length=0,this.max_chain_length=0,this.max_lazy_match=0,this.level=0,this.strategy=0,this.good_match=0,this.nice_match=0,this.dyn_ltree=new c.Buf16(2*w),this.dyn_dtree=new c.Buf16(2*(2*a+1)),this.bl_tree=new c.Buf16(2*(2*o+1)),D(this.dyn_ltree),D(this.dyn_dtree),D(this.bl_tree),this.l_desc=null,this.d_desc=null,this.bl_desc=null,this.bl_count=new c.Buf16(k+1),this.heap=new c.Buf16(2*s+1),D(this.heap),this.heap_len=0,this.heap_max=0,this.depth=new c.Buf16(2*s+1),D(this.depth),this.l_buf=0,this.lit_bufsize=0,this.last_lit=0,this.d_buf=0,this.opt_len=0,this.static_len=0,this.matches=0,this.insert=0,this.bi_buf=0,this.bi_valid=0;}function G(e){var t;return e&&e.state?(e.total_in=e.total_out=0,e.data_type=i,(t=e.state).pending=0,t.pending_out=0,t.wrap<0&&(t.wrap=-t.wrap),t.status=t.wrap?C:E,e.adler=2===t.wrap?0:1,t.last_flush=l,u._tr_init(t),m):R(e,_)}function K(e){var t=G(e);return t===m&&function(e){e.window_size=2*e.w_size,D(e.head),e.max_lazy_match=h[e.level].max_lazy,e.good_match=h[e.level].good_length,e.nice_match=h[e.level].nice_length,e.max_chain_length=h[e.level].max_chain,e.strstart=0,e.block_start=0,e.lookahead=0,e.insert=0,e.match_length=e.prev_length=x-1,e.match_available=0,e.ins_h=0;}(e.state),t}function Y(e,t,r,n,i,s){if(!e)return _;var a=1;if(t===g&&(t=6),n<0?(a=0,n=-n):15<n&&(a=2,n-=16),i<1||y<i||r!==v||n<8||15<n||t<0||9<t||s<0||b<s)return R(e,_);8===n&&(n=9);var o=new H;return (e.state=o).strm=e,o.wrap=a,o.gzhead=null,o.w_bits=n,o.w_size=1<<o.w_bits,o.w_mask=o.w_size-1,o.hash_bits=i+7,o.hash_size=1<<o.hash_bits,o.hash_mask=o.hash_size-1,o.hash_shift=~~((o.hash_bits+x-1)/x),o.window=new c.Buf8(2*o.w_size),o.head=new c.Buf16(o.hash_size),o.prev=new c.Buf16(o.w_size),o.lit_bufsize=1<<i+6,o.pending_buf_size=4*o.lit_bufsize,o.pending_buf=new c.Buf8(o.pending_buf_size),o.d_buf=1*o.lit_bufsize,o.l_buf=3*o.lit_bufsize,o.level=t,o.strategy=s,o.method=r,K(e)}h=[new M(0,0,0,0,function(e,t){var r=65535;for(r>e.pending_buf_size-5&&(r=e.pending_buf_size-5);;){if(e.lookahead<=1){if(j(e),0===e.lookahead&&t===l)return A;if(0===e.lookahead)break}e.strstart+=e.lookahead,e.lookahead=0;var n=e.block_start+r;if((0===e.strstart||e.strstart>=n)&&(e.lookahead=e.strstart-n,e.strstart=n,N(e,false),0===e.strm.avail_out))return A;if(e.strstart-e.block_start>=e.w_size-z&&(N(e,false),0===e.strm.avail_out))return A}return e.insert=0,t===f?(N(e,true),0===e.strm.avail_out?O:B):(e.strstart>e.block_start&&(N(e,false),e.strm.avail_out),A)}),new M(4,4,8,4,Z),new M(4,5,16,8,Z),new M(4,6,32,32,Z),new M(4,4,16,16,W),new M(8,16,32,32,W),new M(8,16,128,128,W),new M(8,32,128,256,W),new M(32,128,258,1024,W),new M(32,258,258,4096,W)],r.deflateInit=function(e,t){return Y(e,t,v,15,8,0)},r.deflateInit2=Y,r.deflateReset=K,r.deflateResetKeep=G,r.deflateSetHeader=function(e,t){return e&&e.state?2!==e.state.wrap?_:(e.state.gzhead=t,m):_},r.deflate=function(e,t){var r,n,i,s;if(!e||!e.state||5<t||t<0)return e?R(e,_):_;if(n=e.state,!e.output||!e.input&&0!==e.avail_in||666===n.status&&t!==f)return R(e,0===e.avail_out?-5:_);if(n.strm=e,r=n.last_flush,n.last_flush=t,n.status===C)if(2===n.wrap)e.adler=0,U(n,31),U(n,139),U(n,8),n.gzhead?(U(n,(n.gzhead.text?1:0)+(n.gzhead.hcrc?2:0)+(n.gzhead.extra?4:0)+(n.gzhead.name?8:0)+(n.gzhead.comment?16:0)),U(n,255&n.gzhead.time),U(n,n.gzhead.time>>8&255),U(n,n.gzhead.time>>16&255),U(n,n.gzhead.time>>24&255),U(n,9===n.level?2:2<=n.strategy||n.level<2?4:0),U(n,255&n.gzhead.os),n.gzhead.extra&&n.gzhead.extra.length&&(U(n,255&n.gzhead.extra.length),U(n,n.gzhead.extra.length>>8&255)),n.gzhead.hcrc&&(e.adler=p(e.adler,n.pending_buf,n.pending,0)),n.gzindex=0,n.status=69):(U(n,0),U(n,0),U(n,0),U(n,0),U(n,0),U(n,9===n.level?2:2<=n.strategy||n.level<2?4:0),U(n,3),n.status=E);else {var a=v+(n.w_bits-8<<4)<<8;a|=(2<=n.strategy||n.level<2?0:n.level<6?1:6===n.level?2:3)<<6,0!==n.strstart&&(a|=32),a+=31-a%31,n.status=E,P(n,a),0!==n.strstart&&(P(n,e.adler>>>16),P(n,65535&e.adler)),e.adler=1;}if(69===n.status)if(n.gzhead.extra){for(i=n.pending;n.gzindex<(65535&n.gzhead.extra.length)&&(n.pending!==n.pending_buf_size||(n.gzhead.hcrc&&n.pending>i&&(e.adler=p(e.adler,n.pending_buf,n.pending-i,i)),F(e),i=n.pending,n.pending!==n.pending_buf_size));)U(n,255&n.gzhead.extra[n.gzindex]),n.gzindex++;n.gzhead.hcrc&&n.pending>i&&(e.adler=p(e.adler,n.pending_buf,n.pending-i,i)),n.gzindex===n.gzhead.extra.length&&(n.gzindex=0,n.status=73);}else n.status=73;if(73===n.status)if(n.gzhead.name){i=n.pending;do{if(n.pending===n.pending_buf_size&&(n.gzhead.hcrc&&n.pending>i&&(e.adler=p(e.adler,n.pending_buf,n.pending-i,i)),F(e),i=n.pending,n.pending===n.pending_buf_size)){s=1;break}s=n.gzindex<n.gzhead.name.length?255&n.gzhead.name.charCodeAt(n.gzindex++):0,U(n,s);}while(0!==s);n.gzhead.hcrc&&n.pending>i&&(e.adler=p(e.adler,n.pending_buf,n.pending-i,i)),0===s&&(n.gzindex=0,n.status=91);}else n.status=91;if(91===n.status)if(n.gzhead.comment){i=n.pending;do{if(n.pending===n.pending_buf_size&&(n.gzhead.hcrc&&n.pending>i&&(e.adler=p(e.adler,n.pending_buf,n.pending-i,i)),F(e),i=n.pending,n.pending===n.pending_buf_size)){s=1;break}s=n.gzindex<n.gzhead.comment.length?255&n.gzhead.comment.charCodeAt(n.gzindex++):0,U(n,s);}while(0!==s);n.gzhead.hcrc&&n.pending>i&&(e.adler=p(e.adler,n.pending_buf,n.pending-i,i)),0===s&&(n.status=103);}else n.status=103;if(103===n.status&&(n.gzhead.hcrc?(n.pending+2>n.pending_buf_size&&F(e),n.pending+2<=n.pending_buf_size&&(U(n,255&e.adler),U(n,e.adler>>8&255),e.adler=0,n.status=E)):n.status=E),0!==n.pending){if(F(e),0===e.avail_out)return n.last_flush=-1,m}else if(0===e.avail_in&&T(t)<=T(r)&&t!==f)return R(e,-5);if(666===n.status&&0!==e.avail_in)return R(e,-5);if(0!==e.avail_in||0!==n.lookahead||t!==l&&666!==n.status){var o=2===n.strategy?function(e,t){for(var r;;){if(0===e.lookahead&&(j(e),0===e.lookahead)){if(t===l)return A;break}if(e.match_length=0,r=u._tr_tally(e,0,e.window[e.strstart]),e.lookahead--,e.strstart++,r&&(N(e,false),0===e.strm.avail_out))return A}return e.insert=0,t===f?(N(e,true),0===e.strm.avail_out?O:B):e.last_lit&&(N(e,false),0===e.strm.avail_out)?A:I}(n,t):3===n.strategy?function(e,t){for(var r,n,i,s,a=e.window;;){if(e.lookahead<=S){if(j(e),e.lookahead<=S&&t===l)return A;if(0===e.lookahead)break}if(e.match_length=0,e.lookahead>=x&&0<e.strstart&&(n=a[i=e.strstart-1])===a[++i]&&n===a[++i]&&n===a[++i]){s=e.strstart+S;do{}while(n===a[++i]&&n===a[++i]&&n===a[++i]&&n===a[++i]&&n===a[++i]&&n===a[++i]&&n===a[++i]&&n===a[++i]&&i<s);e.match_length=S-(s-i),e.match_length>e.lookahead&&(e.match_length=e.lookahead);}if(e.match_length>=x?(r=u._tr_tally(e,1,e.match_length-x),e.lookahead-=e.match_length,e.strstart+=e.match_length,e.match_length=0):(r=u._tr_tally(e,0,e.window[e.strstart]),e.lookahead--,e.strstart++),r&&(N(e,false),0===e.strm.avail_out))return A}return e.insert=0,t===f?(N(e,true),0===e.strm.avail_out?O:B):e.last_lit&&(N(e,false),0===e.strm.avail_out)?A:I}(n,t):h[n.level].func(n,t);if(o!==O&&o!==B||(n.status=666),o===A||o===O)return 0===e.avail_out&&(n.last_flush=-1),m;if(o===I&&(1===t?u._tr_align(n):5!==t&&(u._tr_stored_block(n,0,0,false),3===t&&(D(n.head),0===n.lookahead&&(n.strstart=0,n.block_start=0,n.insert=0))),F(e),0===e.avail_out))return n.last_flush=-1,m}return t!==f?m:n.wrap<=0?1:(2===n.wrap?(U(n,255&e.adler),U(n,e.adler>>8&255),U(n,e.adler>>16&255),U(n,e.adler>>24&255),U(n,255&e.total_in),U(n,e.total_in>>8&255),U(n,e.total_in>>16&255),U(n,e.total_in>>24&255)):(P(n,e.adler>>>16),P(n,65535&e.adler)),F(e),0<n.wrap&&(n.wrap=-n.wrap),0!==n.pending?m:1)},r.deflateEnd=function(e){var t;return e&&e.state?(t=e.state.status)!==C&&69!==t&&73!==t&&91!==t&&103!==t&&t!==E&&666!==t?R(e,_):(e.state=null,t===E?R(e,-3):m):_},r.deflateSetDictionary=function(e,t){var r,n,i,s,a,o,h,u,l=t.length;if(!e||!e.state)return _;if(2===(s=(r=e.state).wrap)||1===s&&r.status!==C||r.lookahead)return _;for(1===s&&(e.adler=d(e.adler,t,l,0)),r.wrap=0,l>=r.w_size&&(0===s&&(D(r.head),r.strstart=0,r.block_start=0,r.insert=0),u=new c.Buf8(r.w_size),c.arraySet(u,t,l-r.w_size,r.w_size,0),t=u,l=r.w_size),a=e.avail_in,o=e.next_in,h=e.input,e.avail_in=l,e.next_in=0,e.input=t,j(r);r.lookahead>=x;){for(n=r.strstart,i=r.lookahead-(x-1);r.ins_h=(r.ins_h<<r.hash_shift^r.window[n+x-1])&r.hash_mask,r.prev[n&r.w_mask]=r.head[r.ins_h],r.head[r.ins_h]=n,n++,--i;);r.strstart=n,r.lookahead=x-1,j(r);}return r.strstart+=r.lookahead,r.block_start=r.strstart,r.insert=r.lookahead,r.lookahead=0,r.match_length=r.prev_length=x-1,r.match_available=0,e.next_in=o,e.input=h,e.avail_in=a,r.wrap=s,m},r.deflateInfo="pako deflate (from Nodeca project)";},{"../utils/common":41,"./adler32":43,"./crc32":45,"./messages":51,"./trees":52}],47:[function(e,t,r){t.exports=function(){this.text=0,this.time=0,this.xflags=0,this.os=0,this.extra=null,this.extra_len=0,this.name="",this.comment="",this.hcrc=0,this.done=false;};},{}],48:[function(e,t,r){t.exports=function(e,t){var r,n,i,s,a,o,h,u,l,f,c,d,p,m,_,g,b,v,y,w,k,x,S,z,C;r=e.state,n=e.next_in,z=e.input,i=n+(e.avail_in-5),s=e.next_out,C=e.output,a=s-(t-e.avail_out),o=s+(e.avail_out-257),h=r.dmax,u=r.wsize,l=r.whave,f=r.wnext,c=r.window,d=r.hold,p=r.bits,m=r.lencode,_=r.distcode,g=(1<<r.lenbits)-1,b=(1<<r.distbits)-1;e:do{p<15&&(d+=z[n++]<<p,p+=8,d+=z[n++]<<p,p+=8),v=m[d&g];t:for(;;){if(d>>>=y=v>>>24,p-=y,0===(y=v>>>16&255))C[s++]=65535&v;else {if(!(16&y)){if(0==(64&y)){v=m[(65535&v)+(d&(1<<y)-1)];continue t}if(32&y){r.mode=12;break e}e.msg="invalid literal/length code",r.mode=30;break e}w=65535&v,(y&=15)&&(p<y&&(d+=z[n++]<<p,p+=8),w+=d&(1<<y)-1,d>>>=y,p-=y),p<15&&(d+=z[n++]<<p,p+=8,d+=z[n++]<<p,p+=8),v=_[d&b];r:for(;;){if(d>>>=y=v>>>24,p-=y,!(16&(y=v>>>16&255))){if(0==(64&y)){v=_[(65535&v)+(d&(1<<y)-1)];continue r}e.msg="invalid distance code",r.mode=30;break e}if(k=65535&v,p<(y&=15)&&(d+=z[n++]<<p,(p+=8)<y&&(d+=z[n++]<<p,p+=8)),h<(k+=d&(1<<y)-1)){e.msg="invalid distance too far back",r.mode=30;break e}if(d>>>=y,p-=y,(y=s-a)<k){if(l<(y=k-y)&&r.sane){e.msg="invalid distance too far back",r.mode=30;break e}if(S=c,(x=0)===f){if(x+=u-y,y<w){for(w-=y;C[s++]=c[x++],--y;);x=s-k,S=C;}}else if(f<y){if(x+=u+f-y,(y-=f)<w){for(w-=y;C[s++]=c[x++],--y;);if(x=0,f<w){for(w-=y=f;C[s++]=c[x++],--y;);x=s-k,S=C;}}}else if(x+=f-y,y<w){for(w-=y;C[s++]=c[x++],--y;);x=s-k,S=C;}for(;2<w;)C[s++]=S[x++],C[s++]=S[x++],C[s++]=S[x++],w-=3;w&&(C[s++]=S[x++],1<w&&(C[s++]=S[x++]));}else {for(x=s-k;C[s++]=C[x++],C[s++]=C[x++],C[s++]=C[x++],2<(w-=3););w&&(C[s++]=C[x++],1<w&&(C[s++]=C[x++]));}break}}break}}while(n<i&&s<o);n-=w=p>>3,d&=(1<<(p-=w<<3))-1,e.next_in=n,e.next_out=s,e.avail_in=n<i?i-n+5:5-(n-i),e.avail_out=s<o?o-s+257:257-(s-o),r.hold=d,r.bits=p;};},{}],49:[function(e,t,r){var I=e("../utils/common"),O=e("./adler32"),B=e("./crc32"),R=e("./inffast"),T=e("./inftrees"),D=1,F=2,N=0,U=-2,P=1,n=852,i=592;function L(e){return (e>>>24&255)+(e>>>8&65280)+((65280&e)<<8)+((255&e)<<24)}function s(){this.mode=0,this.last=false,this.wrap=0,this.havedict=false,this.flags=0,this.dmax=0,this.check=0,this.total=0,this.head=null,this.wbits=0,this.wsize=0,this.whave=0,this.wnext=0,this.window=null,this.hold=0,this.bits=0,this.length=0,this.offset=0,this.extra=0,this.lencode=null,this.distcode=null,this.lenbits=0,this.distbits=0,this.ncode=0,this.nlen=0,this.ndist=0,this.have=0,this.next=null,this.lens=new I.Buf16(320),this.work=new I.Buf16(288),this.lendyn=null,this.distdyn=null,this.sane=0,this.back=0,this.was=0;}function a(e){var t;return e&&e.state?(t=e.state,e.total_in=e.total_out=t.total=0,e.msg="",t.wrap&&(e.adler=1&t.wrap),t.mode=P,t.last=0,t.havedict=0,t.dmax=32768,t.head=null,t.hold=0,t.bits=0,t.lencode=t.lendyn=new I.Buf32(n),t.distcode=t.distdyn=new I.Buf32(i),t.sane=1,t.back=-1,N):U}function o(e){var t;return e&&e.state?((t=e.state).wsize=0,t.whave=0,t.wnext=0,a(e)):U}function h(e,t){var r,n;return e&&e.state?(n=e.state,t<0?(r=0,t=-t):(r=1+(t>>4),t<48&&(t&=15)),t&&(t<8||15<t)?U:(null!==n.window&&n.wbits!==t&&(n.window=null),n.wrap=r,n.wbits=t,o(e))):U}function u(e,t){var r,n;return e?(n=new s,(e.state=n).window=null,(r=h(e,t))!==N&&(e.state=null),r):U}var l,f,c=true;function j(e){if(c){var t;for(l=new I.Buf32(512),f=new I.Buf32(32),t=0;t<144;)e.lens[t++]=8;for(;t<256;)e.lens[t++]=9;for(;t<280;)e.lens[t++]=7;for(;t<288;)e.lens[t++]=8;for(T(D,e.lens,0,288,l,0,e.work,{bits:9}),t=0;t<32;)e.lens[t++]=5;T(F,e.lens,0,32,f,0,e.work,{bits:5}),c=false;}e.lencode=l,e.lenbits=9,e.distcode=f,e.distbits=5;}function Z(e,t,r,n){var i,s=e.state;return null===s.window&&(s.wsize=1<<s.wbits,s.wnext=0,s.whave=0,s.window=new I.Buf8(s.wsize)),n>=s.wsize?(I.arraySet(s.window,t,r-s.wsize,s.wsize,0),s.wnext=0,s.whave=s.wsize):(n<(i=s.wsize-s.wnext)&&(i=n),I.arraySet(s.window,t,r-n,i,s.wnext),(n-=i)?(I.arraySet(s.window,t,r-n,n,0),s.wnext=n,s.whave=s.wsize):(s.wnext+=i,s.wnext===s.wsize&&(s.wnext=0),s.whave<s.wsize&&(s.whave+=i))),0}r.inflateReset=o,r.inflateReset2=h,r.inflateResetKeep=a,r.inflateInit=function(e){return u(e,15)},r.inflateInit2=u,r.inflate=function(e,t){var r,n,i,s,a,o,h,u,l,f,c,d,p,m,_,g,b,v,y,w,k,x,S,z,C=0,E=new I.Buf8(4),A=[16,17,18,0,8,7,9,6,10,5,11,4,12,3,13,2,14,1,15];if(!e||!e.state||!e.output||!e.input&&0!==e.avail_in)return U;12===(r=e.state).mode&&(r.mode=13),a=e.next_out,i=e.output,h=e.avail_out,s=e.next_in,n=e.input,o=e.avail_in,u=r.hold,l=r.bits,f=o,c=h,x=N;e:for(;;)switch(r.mode){case P:if(0===r.wrap){r.mode=13;break}for(;l<16;){if(0===o)break e;o--,u+=n[s++]<<l,l+=8;}if(2&r.wrap&&35615===u){E[r.check=0]=255&u,E[1]=u>>>8&255,r.check=B(r.check,E,2,0),l=u=0,r.mode=2;break}if(r.flags=0,r.head&&(r.head.done=false),!(1&r.wrap)||(((255&u)<<8)+(u>>8))%31){e.msg="incorrect header check",r.mode=30;break}if(8!=(15&u)){e.msg="unknown compression method",r.mode=30;break}if(l-=4,k=8+(15&(u>>>=4)),0===r.wbits)r.wbits=k;else if(k>r.wbits){e.msg="invalid window size",r.mode=30;break}r.dmax=1<<k,e.adler=r.check=1,r.mode=512&u?10:12,l=u=0;break;case 2:for(;l<16;){if(0===o)break e;o--,u+=n[s++]<<l,l+=8;}if(r.flags=u,8!=(255&r.flags)){e.msg="unknown compression method",r.mode=30;break}if(57344&r.flags){e.msg="unknown header flags set",r.mode=30;break}r.head&&(r.head.text=u>>8&1),512&r.flags&&(E[0]=255&u,E[1]=u>>>8&255,r.check=B(r.check,E,2,0)),l=u=0,r.mode=3;case 3:for(;l<32;){if(0===o)break e;o--,u+=n[s++]<<l,l+=8;}r.head&&(r.head.time=u),512&r.flags&&(E[0]=255&u,E[1]=u>>>8&255,E[2]=u>>>16&255,E[3]=u>>>24&255,r.check=B(r.check,E,4,0)),l=u=0,r.mode=4;case 4:for(;l<16;){if(0===o)break e;o--,u+=n[s++]<<l,l+=8;}r.head&&(r.head.xflags=255&u,r.head.os=u>>8),512&r.flags&&(E[0]=255&u,E[1]=u>>>8&255,r.check=B(r.check,E,2,0)),l=u=0,r.mode=5;case 5:if(1024&r.flags){for(;l<16;){if(0===o)break e;o--,u+=n[s++]<<l,l+=8;}r.length=u,r.head&&(r.head.extra_len=u),512&r.flags&&(E[0]=255&u,E[1]=u>>>8&255,r.check=B(r.check,E,2,0)),l=u=0;}else r.head&&(r.head.extra=null);r.mode=6;case 6:if(1024&r.flags&&(o<(d=r.length)&&(d=o),d&&(r.head&&(k=r.head.extra_len-r.length,r.head.extra||(r.head.extra=new Array(r.head.extra_len)),I.arraySet(r.head.extra,n,s,d,k)),512&r.flags&&(r.check=B(r.check,n,d,s)),o-=d,s+=d,r.length-=d),r.length))break e;r.length=0,r.mode=7;case 7:if(2048&r.flags){if(0===o)break e;for(d=0;k=n[s+d++],r.head&&k&&r.length<65536&&(r.head.name+=String.fromCharCode(k)),k&&d<o;);if(512&r.flags&&(r.check=B(r.check,n,d,s)),o-=d,s+=d,k)break e}else r.head&&(r.head.name=null);r.length=0,r.mode=8;case 8:if(4096&r.flags){if(0===o)break e;for(d=0;k=n[s+d++],r.head&&k&&r.length<65536&&(r.head.comment+=String.fromCharCode(k)),k&&d<o;);if(512&r.flags&&(r.check=B(r.check,n,d,s)),o-=d,s+=d,k)break e}else r.head&&(r.head.comment=null);r.mode=9;case 9:if(512&r.flags){for(;l<16;){if(0===o)break e;o--,u+=n[s++]<<l,l+=8;}if(u!==(65535&r.check)){e.msg="header crc mismatch",r.mode=30;break}l=u=0;}r.head&&(r.head.hcrc=r.flags>>9&1,r.head.done=true),e.adler=r.check=0,r.mode=12;break;case 10:for(;l<32;){if(0===o)break e;o--,u+=n[s++]<<l,l+=8;}e.adler=r.check=L(u),l=u=0,r.mode=11;case 11:if(0===r.havedict)return e.next_out=a,e.avail_out=h,e.next_in=s,e.avail_in=o,r.hold=u,r.bits=l,2;e.adler=r.check=1,r.mode=12;case 12:if(5===t||6===t)break e;case 13:if(r.last){u>>>=7&l,l-=7&l,r.mode=27;break}for(;l<3;){if(0===o)break e;o--,u+=n[s++]<<l,l+=8;}switch(r.last=1&u,l-=1,3&(u>>>=1)){case 0:r.mode=14;break;case 1:if(j(r),r.mode=20,6!==t)break;u>>>=2,l-=2;break e;case 2:r.mode=17;break;case 3:e.msg="invalid block type",r.mode=30;}u>>>=2,l-=2;break;case 14:for(u>>>=7&l,l-=7&l;l<32;){if(0===o)break e;o--,u+=n[s++]<<l,l+=8;}if((65535&u)!=(u>>>16^65535)){e.msg="invalid stored block lengths",r.mode=30;break}if(r.length=65535&u,l=u=0,r.mode=15,6===t)break e;case 15:r.mode=16;case 16:if(d=r.length){if(o<d&&(d=o),h<d&&(d=h),0===d)break e;I.arraySet(i,n,s,d,a),o-=d,s+=d,h-=d,a+=d,r.length-=d;break}r.mode=12;break;case 17:for(;l<14;){if(0===o)break e;o--,u+=n[s++]<<l,l+=8;}if(r.nlen=257+(31&u),u>>>=5,l-=5,r.ndist=1+(31&u),u>>>=5,l-=5,r.ncode=4+(15&u),u>>>=4,l-=4,286<r.nlen||30<r.ndist){e.msg="too many length or distance symbols",r.mode=30;break}r.have=0,r.mode=18;case 18:for(;r.have<r.ncode;){for(;l<3;){if(0===o)break e;o--,u+=n[s++]<<l,l+=8;}r.lens[A[r.have++]]=7&u,u>>>=3,l-=3;}for(;r.have<19;)r.lens[A[r.have++]]=0;if(r.lencode=r.lendyn,r.lenbits=7,S={bits:r.lenbits},x=T(0,r.lens,0,19,r.lencode,0,r.work,S),r.lenbits=S.bits,x){e.msg="invalid code lengths set",r.mode=30;break}r.have=0,r.mode=19;case 19:for(;r.have<r.nlen+r.ndist;){for(;g=(C=r.lencode[u&(1<<r.lenbits)-1])>>>16&255,b=65535&C,!((_=C>>>24)<=l);){if(0===o)break e;o--,u+=n[s++]<<l,l+=8;}if(b<16)u>>>=_,l-=_,r.lens[r.have++]=b;else {if(16===b){for(z=_+2;l<z;){if(0===o)break e;o--,u+=n[s++]<<l,l+=8;}if(u>>>=_,l-=_,0===r.have){e.msg="invalid bit length repeat",r.mode=30;break}k=r.lens[r.have-1],d=3+(3&u),u>>>=2,l-=2;}else if(17===b){for(z=_+3;l<z;){if(0===o)break e;o--,u+=n[s++]<<l,l+=8;}l-=_,k=0,d=3+(7&(u>>>=_)),u>>>=3,l-=3;}else {for(z=_+7;l<z;){if(0===o)break e;o--,u+=n[s++]<<l,l+=8;}l-=_,k=0,d=11+(127&(u>>>=_)),u>>>=7,l-=7;}if(r.have+d>r.nlen+r.ndist){e.msg="invalid bit length repeat",r.mode=30;break}for(;d--;)r.lens[r.have++]=k;}}if(30===r.mode)break;if(0===r.lens[256]){e.msg="invalid code -- missing end-of-block",r.mode=30;break}if(r.lenbits=9,S={bits:r.lenbits},x=T(D,r.lens,0,r.nlen,r.lencode,0,r.work,S),r.lenbits=S.bits,x){e.msg="invalid literal/lengths set",r.mode=30;break}if(r.distbits=6,r.distcode=r.distdyn,S={bits:r.distbits},x=T(F,r.lens,r.nlen,r.ndist,r.distcode,0,r.work,S),r.distbits=S.bits,x){e.msg="invalid distances set",r.mode=30;break}if(r.mode=20,6===t)break e;case 20:r.mode=21;case 21:if(6<=o&&258<=h){e.next_out=a,e.avail_out=h,e.next_in=s,e.avail_in=o,r.hold=u,r.bits=l,R(e,c),a=e.next_out,i=e.output,h=e.avail_out,s=e.next_in,n=e.input,o=e.avail_in,u=r.hold,l=r.bits,12===r.mode&&(r.back=-1);break}for(r.back=0;g=(C=r.lencode[u&(1<<r.lenbits)-1])>>>16&255,b=65535&C,!((_=C>>>24)<=l);){if(0===o)break e;o--,u+=n[s++]<<l,l+=8;}if(g&&0==(240&g)){for(v=_,y=g,w=b;g=(C=r.lencode[w+((u&(1<<v+y)-1)>>v)])>>>16&255,b=65535&C,!(v+(_=C>>>24)<=l);){if(0===o)break e;o--,u+=n[s++]<<l,l+=8;}u>>>=v,l-=v,r.back+=v;}if(u>>>=_,l-=_,r.back+=_,r.length=b,0===g){r.mode=26;break}if(32&g){r.back=-1,r.mode=12;break}if(64&g){e.msg="invalid literal/length code",r.mode=30;break}r.extra=15&g,r.mode=22;case 22:if(r.extra){for(z=r.extra;l<z;){if(0===o)break e;o--,u+=n[s++]<<l,l+=8;}r.length+=u&(1<<r.extra)-1,u>>>=r.extra,l-=r.extra,r.back+=r.extra;}r.was=r.length,r.mode=23;case 23:for(;g=(C=r.distcode[u&(1<<r.distbits)-1])>>>16&255,b=65535&C,!((_=C>>>24)<=l);){if(0===o)break e;o--,u+=n[s++]<<l,l+=8;}if(0==(240&g)){for(v=_,y=g,w=b;g=(C=r.distcode[w+((u&(1<<v+y)-1)>>v)])>>>16&255,b=65535&C,!(v+(_=C>>>24)<=l);){if(0===o)break e;o--,u+=n[s++]<<l,l+=8;}u>>>=v,l-=v,r.back+=v;}if(u>>>=_,l-=_,r.back+=_,64&g){e.msg="invalid distance code",r.mode=30;break}r.offset=b,r.extra=15&g,r.mode=24;case 24:if(r.extra){for(z=r.extra;l<z;){if(0===o)break e;o--,u+=n[s++]<<l,l+=8;}r.offset+=u&(1<<r.extra)-1,u>>>=r.extra,l-=r.extra,r.back+=r.extra;}if(r.offset>r.dmax){e.msg="invalid distance too far back",r.mode=30;break}r.mode=25;case 25:if(0===h)break e;if(d=c-h,r.offset>d){if((d=r.offset-d)>r.whave&&r.sane){e.msg="invalid distance too far back",r.mode=30;break}p=d>r.wnext?(d-=r.wnext,r.wsize-d):r.wnext-d,d>r.length&&(d=r.length),m=r.window;}else m=i,p=a-r.offset,d=r.length;for(h<d&&(d=h),h-=d,r.length-=d;i[a++]=m[p++],--d;);0===r.length&&(r.mode=21);break;case 26:if(0===h)break e;i[a++]=r.length,h--,r.mode=21;break;case 27:if(r.wrap){for(;l<32;){if(0===o)break e;o--,u|=n[s++]<<l,l+=8;}if(c-=h,e.total_out+=c,r.total+=c,c&&(e.adler=r.check=r.flags?B(r.check,i,c,a-c):O(r.check,i,c,a-c)),c=h,(r.flags?u:L(u))!==r.check){e.msg="incorrect data check",r.mode=30;break}l=u=0;}r.mode=28;case 28:if(r.wrap&&r.flags){for(;l<32;){if(0===o)break e;o--,u+=n[s++]<<l,l+=8;}if(u!==(4294967295&r.total)){e.msg="incorrect length check",r.mode=30;break}l=u=0;}r.mode=29;case 29:x=1;break e;case 30:x=-3;break e;case 31:return  -4;case 32:default:return U}return e.next_out=a,e.avail_out=h,e.next_in=s,e.avail_in=o,r.hold=u,r.bits=l,(r.wsize||c!==e.avail_out&&r.mode<30&&(r.mode<27||4!==t))&&Z(e,e.output,e.next_out,c-e.avail_out)?(r.mode=31,-4):(f-=e.avail_in,c-=e.avail_out,e.total_in+=f,e.total_out+=c,r.total+=c,r.wrap&&c&&(e.adler=r.check=r.flags?B(r.check,i,c,e.next_out-c):O(r.check,i,c,e.next_out-c)),e.data_type=r.bits+(r.last?64:0)+(12===r.mode?128:0)+(20===r.mode||15===r.mode?256:0),(0==f&&0===c||4===t)&&x===N&&(x=-5),x)},r.inflateEnd=function(e){if(!e||!e.state)return U;var t=e.state;return t.window&&(t.window=null),e.state=null,N},r.inflateGetHeader=function(e,t){var r;return e&&e.state?0==(2&(r=e.state).wrap)?U:((r.head=t).done=false,N):U},r.inflateSetDictionary=function(e,t){var r,n=t.length;return e&&e.state?0!==(r=e.state).wrap&&11!==r.mode?U:11===r.mode&&O(1,t,n,0)!==r.check?-3:Z(e,t,n,n)?(r.mode=31,-4):(r.havedict=1,N):U},r.inflateInfo="pako inflate (from Nodeca project)";},{"../utils/common":41,"./adler32":43,"./crc32":45,"./inffast":48,"./inftrees":50}],50:[function(e,t,r){var D=e("../utils/common"),F=[3,4,5,6,7,8,9,10,11,13,15,17,19,23,27,31,35,43,51,59,67,83,99,115,131,163,195,227,258,0,0],N=[16,16,16,16,16,16,16,16,17,17,17,17,18,18,18,18,19,19,19,19,20,20,20,20,21,21,21,21,16,72,78],U=[1,2,3,4,5,7,9,13,17,25,33,49,65,97,129,193,257,385,513,769,1025,1537,2049,3073,4097,6145,8193,12289,16385,24577,0,0],P=[16,16,16,16,17,17,18,18,19,19,20,20,21,21,22,22,23,23,24,24,25,25,26,26,27,27,28,28,29,29,64,64];t.exports=function(e,t,r,n,i,s,a,o){var h,u,l,f,c,d,p,m,_,g=o.bits,b=0,v=0,y=0,w=0,k=0,x=0,S=0,z=0,C=0,E=0,A=null,I=0,O=new D.Buf16(16),B=new D.Buf16(16),R=null,T=0;for(b=0;b<=15;b++)O[b]=0;for(v=0;v<n;v++)O[t[r+v]]++;for(k=g,w=15;1<=w&&0===O[w];w--);if(w<k&&(k=w),0===w)return i[s++]=20971520,i[s++]=20971520,o.bits=1,0;for(y=1;y<w&&0===O[y];y++);for(k<y&&(k=y),b=z=1;b<=15;b++)if(z<<=1,(z-=O[b])<0)return  -1;if(0<z&&(0===e||1!==w))return  -1;for(B[1]=0,b=1;b<15;b++)B[b+1]=B[b]+O[b];for(v=0;v<n;v++)0!==t[r+v]&&(a[B[t[r+v]]++]=v);if(d=0===e?(A=R=a,19):1===e?(A=F,I-=257,R=N,T-=257,256):(A=U,R=P,-1),b=y,c=s,S=v=E=0,l=-1,f=(C=1<<(x=k))-1,1===e&&852<C||2===e&&592<C)return 1;for(;;){for(p=b-S,_=a[v]<d?(m=0,a[v]):a[v]>d?(m=R[T+a[v]],A[I+a[v]]):(m=96,0),h=1<<b-S,y=u=1<<x;i[c+(E>>S)+(u-=h)]=p<<24|m<<16|_|0,0!==u;);for(h=1<<b-1;E&h;)h>>=1;if(0!==h?(E&=h-1,E+=h):E=0,v++,0==--O[b]){if(b===w)break;b=t[r+a[v]];}if(k<b&&(E&f)!==l){for(0===S&&(S=k),c+=y,z=1<<(x=b-S);x+S<w&&!((z-=O[x+S])<=0);)x++,z<<=1;if(C+=1<<x,1===e&&852<C||2===e&&592<C)return 1;i[l=E&f]=k<<24|x<<16|c-s|0;}}return 0!==E&&(i[c+E]=b-S<<24|64<<16|0),o.bits=k,0};},{"../utils/common":41}],51:[function(e,t,r){t.exports={2:"need dictionary",1:"stream end",0:"","-1":"file error","-2":"stream error","-3":"data error","-4":"insufficient memory","-5":"buffer error","-6":"incompatible version"};},{}],52:[function(e,t,r){var i=e("../utils/common"),o=0,h=1;function n(e){for(var t=e.length;0<=--t;)e[t]=0;}var s=0,a=29,u=256,l=u+1+a,f=30,c=19,_=2*l+1,g=15,d=16,p=7,m=256,b=16,v=17,y=18,w=[0,0,0,0,0,0,0,0,1,1,1,1,2,2,2,2,3,3,3,3,4,4,4,4,5,5,5,5,0],k=[0,0,0,0,1,1,2,2,3,3,4,4,5,5,6,6,7,7,8,8,9,9,10,10,11,11,12,12,13,13],x=[0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,2,3,7],S=[16,17,18,0,8,7,9,6,10,5,11,4,12,3,13,2,14,1,15],z=new Array(2*(l+2));n(z);var C=new Array(2*f);n(C);var E=new Array(512);n(E);var A=new Array(256);n(A);var I=new Array(a);n(I);var O,B,R,T=new Array(f);function D(e,t,r,n,i){this.static_tree=e,this.extra_bits=t,this.extra_base=r,this.elems=n,this.max_length=i,this.has_stree=e&&e.length;}function F(e,t){this.dyn_tree=e,this.max_code=0,this.stat_desc=t;}function N(e){return e<256?E[e]:E[256+(e>>>7)]}function U(e,t){e.pending_buf[e.pending++]=255&t,e.pending_buf[e.pending++]=t>>>8&255;}function P(e,t,r){e.bi_valid>d-r?(e.bi_buf|=t<<e.bi_valid&65535,U(e,e.bi_buf),e.bi_buf=t>>d-e.bi_valid,e.bi_valid+=r-d):(e.bi_buf|=t<<e.bi_valid&65535,e.bi_valid+=r);}function L(e,t,r){P(e,r[2*t],r[2*t+1]);}function j(e,t){for(var r=0;r|=1&e,e>>>=1,r<<=1,0<--t;);return r>>>1}function Z(e,t,r){var n,i,s=new Array(g+1),a=0;for(n=1;n<=g;n++)s[n]=a=a+r[n-1]<<1;for(i=0;i<=t;i++){var o=e[2*i+1];0!==o&&(e[2*i]=j(s[o]++,o));}}function W(e){var t;for(t=0;t<l;t++)e.dyn_ltree[2*t]=0;for(t=0;t<f;t++)e.dyn_dtree[2*t]=0;for(t=0;t<c;t++)e.bl_tree[2*t]=0;e.dyn_ltree[2*m]=1,e.opt_len=e.static_len=0,e.last_lit=e.matches=0;}function M(e){8<e.bi_valid?U(e,e.bi_buf):0<e.bi_valid&&(e.pending_buf[e.pending++]=e.bi_buf),e.bi_buf=0,e.bi_valid=0;}function H(e,t,r,n){var i=2*t,s=2*r;return e[i]<e[s]||e[i]===e[s]&&n[t]<=n[r]}function G(e,t,r){for(var n=e.heap[r],i=r<<1;i<=e.heap_len&&(i<e.heap_len&&H(t,e.heap[i+1],e.heap[i],e.depth)&&i++,!H(t,n,e.heap[i],e.depth));)e.heap[r]=e.heap[i],r=i,i<<=1;e.heap[r]=n;}function K(e,t,r){var n,i,s,a,o=0;if(0!==e.last_lit)for(;n=e.pending_buf[e.d_buf+2*o]<<8|e.pending_buf[e.d_buf+2*o+1],i=e.pending_buf[e.l_buf+o],o++,0===n?L(e,i,t):(L(e,(s=A[i])+u+1,t),0!==(a=w[s])&&P(e,i-=I[s],a),L(e,s=N(--n),r),0!==(a=k[s])&&P(e,n-=T[s],a)),o<e.last_lit;);L(e,m,t);}function Y(e,t){var r,n,i,s=t.dyn_tree,a=t.stat_desc.static_tree,o=t.stat_desc.has_stree,h=t.stat_desc.elems,u=-1;for(e.heap_len=0,e.heap_max=_,r=0;r<h;r++)0!==s[2*r]?(e.heap[++e.heap_len]=u=r,e.depth[r]=0):s[2*r+1]=0;for(;e.heap_len<2;)s[2*(i=e.heap[++e.heap_len]=u<2?++u:0)]=1,e.depth[i]=0,e.opt_len--,o&&(e.static_len-=a[2*i+1]);for(t.max_code=u,r=e.heap_len>>1;1<=r;r--)G(e,s,r);for(i=h;r=e.heap[1],e.heap[1]=e.heap[e.heap_len--],G(e,s,1),n=e.heap[1],e.heap[--e.heap_max]=r,e.heap[--e.heap_max]=n,s[2*i]=s[2*r]+s[2*n],e.depth[i]=(e.depth[r]>=e.depth[n]?e.depth[r]:e.depth[n])+1,s[2*r+1]=s[2*n+1]=i,e.heap[1]=i++,G(e,s,1),2<=e.heap_len;);e.heap[--e.heap_max]=e.heap[1],function(e,t){var r,n,i,s,a,o,h=t.dyn_tree,u=t.max_code,l=t.stat_desc.static_tree,f=t.stat_desc.has_stree,c=t.stat_desc.extra_bits,d=t.stat_desc.extra_base,p=t.stat_desc.max_length,m=0;for(s=0;s<=g;s++)e.bl_count[s]=0;for(h[2*e.heap[e.heap_max]+1]=0,r=e.heap_max+1;r<_;r++)p<(s=h[2*h[2*(n=e.heap[r])+1]+1]+1)&&(s=p,m++),h[2*n+1]=s,u<n||(e.bl_count[s]++,a=0,d<=n&&(a=c[n-d]),o=h[2*n],e.opt_len+=o*(s+a),f&&(e.static_len+=o*(l[2*n+1]+a)));if(0!==m){do{for(s=p-1;0===e.bl_count[s];)s--;e.bl_count[s]--,e.bl_count[s+1]+=2,e.bl_count[p]--,m-=2;}while(0<m);for(s=p;0!==s;s--)for(n=e.bl_count[s];0!==n;)u<(i=e.heap[--r])||(h[2*i+1]!==s&&(e.opt_len+=(s-h[2*i+1])*h[2*i],h[2*i+1]=s),n--);}}(e,t),Z(s,u,e.bl_count);}function X(e,t,r){var n,i,s=-1,a=t[1],o=0,h=7,u=4;for(0===a&&(h=138,u=3),t[2*(r+1)+1]=65535,n=0;n<=r;n++)i=a,a=t[2*(n+1)+1],++o<h&&i===a||(o<u?e.bl_tree[2*i]+=o:0!==i?(i!==s&&e.bl_tree[2*i]++,e.bl_tree[2*b]++):o<=10?e.bl_tree[2*v]++:e.bl_tree[2*y]++,s=i,u=(o=0)===a?(h=138,3):i===a?(h=6,3):(h=7,4));}function V(e,t,r){var n,i,s=-1,a=t[1],o=0,h=7,u=4;for(0===a&&(h=138,u=3),n=0;n<=r;n++)if(i=a,a=t[2*(n+1)+1],!(++o<h&&i===a)){if(o<u)for(;L(e,i,e.bl_tree),0!=--o;);else 0!==i?(i!==s&&(L(e,i,e.bl_tree),o--),L(e,b,e.bl_tree),P(e,o-3,2)):o<=10?(L(e,v,e.bl_tree),P(e,o-3,3)):(L(e,y,e.bl_tree),P(e,o-11,7));s=i,u=(o=0)===a?(h=138,3):i===a?(h=6,3):(h=7,4);}}n(T);var q=false;function J(e,t,r,n){P(e,(s<<1)+(n?1:0),3),function(e,t,r,n){M(e),(U(e,r),U(e,~r)),i.arraySet(e.pending_buf,e.window,t,r,e.pending),e.pending+=r;}(e,t,r);}r._tr_init=function(e){q||(function(){var e,t,r,n,i,s=new Array(g+1);for(n=r=0;n<a-1;n++)for(I[n]=r,e=0;e<1<<w[n];e++)A[r++]=n;for(A[r-1]=n,n=i=0;n<16;n++)for(T[n]=i,e=0;e<1<<k[n];e++)E[i++]=n;for(i>>=7;n<f;n++)for(T[n]=i<<7,e=0;e<1<<k[n]-7;e++)E[256+i++]=n;for(t=0;t<=g;t++)s[t]=0;for(e=0;e<=143;)z[2*e+1]=8,e++,s[8]++;for(;e<=255;)z[2*e+1]=9,e++,s[9]++;for(;e<=279;)z[2*e+1]=7,e++,s[7]++;for(;e<=287;)z[2*e+1]=8,e++,s[8]++;for(Z(z,l+1,s),e=0;e<f;e++)C[2*e+1]=5,C[2*e]=j(e,5);O=new D(z,w,u+1,l,g),B=new D(C,k,0,f,g),R=new D(new Array(0),x,0,c,p);}(),q=true),e.l_desc=new F(e.dyn_ltree,O),e.d_desc=new F(e.dyn_dtree,B),e.bl_desc=new F(e.bl_tree,R),e.bi_buf=0,e.bi_valid=0,W(e);},r._tr_stored_block=J,r._tr_flush_block=function(e,t,r,n){var i,s,a=0;0<e.level?(2===e.strm.data_type&&(e.strm.data_type=function(e){var t,r=4093624447;for(t=0;t<=31;t++,r>>>=1)if(1&r&&0!==e.dyn_ltree[2*t])return o;if(0!==e.dyn_ltree[18]||0!==e.dyn_ltree[20]||0!==e.dyn_ltree[26])return h;for(t=32;t<u;t++)if(0!==e.dyn_ltree[2*t])return h;return o}(e)),Y(e,e.l_desc),Y(e,e.d_desc),a=function(e){var t;for(X(e,e.dyn_ltree,e.l_desc.max_code),X(e,e.dyn_dtree,e.d_desc.max_code),Y(e,e.bl_desc),t=c-1;3<=t&&0===e.bl_tree[2*S[t]+1];t--);return e.opt_len+=3*(t+1)+5+5+4,t}(e),i=e.opt_len+3+7>>>3,(s=e.static_len+3+7>>>3)<=i&&(i=s)):i=s=r+5,r+4<=i&&-1!==t?J(e,t,r,n):4===e.strategy||s===i?(P(e,2+(n?1:0),3),K(e,z,C)):(P(e,4+(n?1:0),3),function(e,t,r,n){var i;for(P(e,t-257,5),P(e,r-1,5),P(e,n-4,4),i=0;i<n;i++)P(e,e.bl_tree[2*S[i]+1],3);V(e,e.dyn_ltree,t-1),V(e,e.dyn_dtree,r-1);}(e,e.l_desc.max_code+1,e.d_desc.max_code+1,a+1),K(e,e.dyn_ltree,e.dyn_dtree)),W(e),n&&M(e);},r._tr_tally=function(e,t,r){return e.pending_buf[e.d_buf+2*e.last_lit]=t>>>8&255,e.pending_buf[e.d_buf+2*e.last_lit+1]=255&t,e.pending_buf[e.l_buf+e.last_lit]=255&r,e.last_lit++,0===t?e.dyn_ltree[2*r]++:(e.matches++,t--,e.dyn_ltree[2*(A[r]+u+1)]++,e.dyn_dtree[2*N(t)]++),e.last_lit===e.lit_bufsize-1},r._tr_align=function(e){P(e,2,3),L(e,m,z),function(e){16===e.bi_valid?(U(e,e.bi_buf),e.bi_buf=0,e.bi_valid=0):8<=e.bi_valid&&(e.pending_buf[e.pending++]=255&e.bi_buf,e.bi_buf>>=8,e.bi_valid-=8);}(e);};},{"../utils/common":41}],53:[function(e,t,r){t.exports=function(){this.input=null,this.next_in=0,this.avail_in=0,this.total_in=0,this.output=null,this.next_out=0,this.avail_out=0,this.total_out=0,this.msg="",this.state=null,this.data_type=2,this.adler=0;};},{}],54:[function(e,t,r){(function(e){!function(r,n){if(!r.setImmediate){var i,s,t,a,o=1,h={},u=false,l=r.document,e=Object.getPrototypeOf&&Object.getPrototypeOf(r);e=e&&e.setTimeout?e:r,i="[object process]"==={}.toString.call(r.process)?function(e){process.nextTick(function(){c(e);});}:function(){if(r.postMessage&&!r.importScripts){var e=true,t=r.onmessage;return r.onmessage=function(){e=false;},r.postMessage("","*"),r.onmessage=t,e}}()?(a="setImmediate$"+Math.random()+"$",r.addEventListener?r.addEventListener("message",d,false):r.attachEvent("onmessage",d),function(e){r.postMessage(a+e,"*");}):r.MessageChannel?((t=new MessageChannel).port1.onmessage=function(e){c(e.data);},function(e){t.port2.postMessage(e);}):l&&"onreadystatechange"in l.createElement("script")?(s=l.documentElement,function(e){var t=l.createElement("script");t.onreadystatechange=function(){c(e),t.onreadystatechange=null,s.removeChild(t),t=null;},s.appendChild(t);}):function(e){setTimeout(c,0,e);},e.setImmediate=function(e){"function"!=typeof e&&(e=new Function(""+e));for(var t=new Array(arguments.length-1),r=0;r<t.length;r++)t[r]=arguments[r+1];var n={callback:e,args:t};return h[o]=n,i(o),o++},e.clearImmediate=f;}function f(e){delete h[e];}function c(e){if(u)setTimeout(c,0,e);else {var t=h[e];if(t){u=true;try{!function(e){var t=e.callback,r=e.args;switch(r.length){case 0:t();break;case 1:t(r[0]);break;case 2:t(r[0],r[1]);break;case 3:t(r[0],r[1],r[2]);break;default:t.apply(n,r);}}(t);}finally{f(e),u=false;}}}}function d(e){e.source===r&&"string"==typeof e.data&&0===e.data.indexOf(a)&&c(+e.data.slice(a.length));}}("undefined"==typeof self?void 0===e?this:e:self);}).call(this,"undefined"!=typeof commonjsGlobal?commonjsGlobal:"undefined"!=typeof self?self:"undefined"!=typeof window?window:{});},{}]},{},[10])(10)}); 
} (jszip_min));

var jszip_minExports = jszip_min.exports;
var JSZip = /*@__PURE__*/getDefaultExportFromCjs(jszip_minExports);

const SLIDE_FACTOR$1 = 96 / 914400;
const FONT_SIZE_FACTOR = 4 / 3.2;
const RTL_LANGS_ARRAY = [
    "he-IL", "ar-AE", "ar-SA", "ar-EG", "ar-IQ", "ar-JO", "ar-KW", "ar-LB", "ar-LY",
    "ar-MA", "ar-OM", "ar-PS", "ar-QA", "ar-SD", "ar-SY", "ar-TN", "ar-YE",
    "fa-IR", "fa-AF",
    "ur-PK", "ur-IN",
    "dv-MV",
    "szq-DZ",
    "ps-AF", "ps-PK",
    "ug-CN",
    "kk-KZ", "ky-KG", "uz-UZ", "tg-TJ",
    "yi-DE", "jpr-IL", "jrb-IL"
];
const DINGBAT_UNICODE = [
    { f: "Webdings", code: "33", unicode: "128375" },
    { f: "Webdings", code: "34", unicode: "128376" },
    { f: "Webdings", code: "35", unicode: "128370" },
    { f: "Webdings", code: "36", unicode: "128374" },
    { f: "Webdings", code: "37", unicode: "127942" },
    { f: "Webdings", code: "38", unicode: "127894" },
    { f: "Webdings", code: "39", unicode: "128391" },
    { f: "Webdings", code: "40", unicode: "128488" },
    { f: "Webdings", code: "41", unicode: "128489" },
    { f: "Webdings", code: "42", unicode: "128496" },
    { f: "Webdings", code: "43", unicode: "128497" },
    { f: "Webdings", code: "44", unicode: "127798" },
    { f: "Webdings", code: "45", unicode: "127895" },
    { f: "Webdings", code: "46", unicode: "128638" },
    { f: "Webdings", code: "47", unicode: "128636" },
    { f: "Webdings", code: "48", unicode: "128469" },
    { f: "Webdings", code: "49", unicode: "128470" },
    { f: "Webdings", code: "50", unicode: "128471" },
    { f: "Webdings", code: "51", unicode: "9204" },
    { f: "Webdings", code: "52", unicode: "9205" },
    { f: "Webdings", code: "53", unicode: "9206" },
    { f: "Webdings", code: "54", unicode: "9207" },
    { f: "Webdings", code: "55", unicode: "9194" },
    { f: "Webdings", code: "56", unicode: "9193" },
    { f: "Webdings", code: "57", unicode: "9198" },
    { f: "Webdings", code: "58", unicode: "9197" },
    { f: "Webdings", code: "59", unicode: "9208" },
    { f: "Webdings", code: "60", unicode: "9209" },
    { f: "Webdings", code: "61", unicode: "9210" },
    { f: "Webdings", code: "62", unicode: "128474" },
    { f: "Webdings", code: "63", unicode: "128499" },
    { f: "Webdings", code: "64", unicode: "128736" },
    { f: "Webdings", code: "65", unicode: "127959" },
    { f: "Webdings", code: "66", unicode: "127960" },
    { f: "Webdings", code: "67", unicode: "127961" },
    { f: "Webdings", code: "68", unicode: "127962" },
    { f: "Webdings", code: "69", unicode: "127964" },
    { f: "Webdings", code: "70", unicode: "127981" },
    { f: "Webdings", code: "71", unicode: "127963" },
    { f: "Webdings", code: "72", unicode: "127968" },
    { f: "Webdings", code: "73", unicode: "127958" },
    { f: "Webdings", code: "74", unicode: "127965" },
    { f: "Webdings", code: "75", unicode: "128739" },
    { f: "Webdings", code: "76", unicode: "128269" },
    { f: "Webdings", code: "77", unicode: "127956" },
    { f: "Webdings", code: "78", unicode: "128065" },
    { f: "Webdings", code: "79", unicode: "128066" },
    { f: "Webdings", code: "80", unicode: "127966" },
    { f: "Webdings", code: "81", unicode: "127957" },
    { f: "Webdings", code: "82", unicode: "128740" },
    { f: "Webdings", code: "83", unicode: "127967" },
    { f: "Webdings", code: "84", unicode: "128755" },
    { f: "Webdings", code: "85", unicode: "128364" },
    { f: "Webdings", code: "86", unicode: "128363" },
    { f: "Webdings", code: "87", unicode: "128360" },
    { f: "Webdings", code: "88", unicode: "128264" },
    { f: "Webdings", code: "89", unicode: "127892" },
    { f: "Webdings", code: "90", unicode: "127893" },
    { f: "Webdings", code: "91", unicode: "128492" },
    { f: "Webdings", code: "92", unicode: "128637" },
    { f: "Webdings", code: "93", unicode: "128493" },
    { f: "Webdings", code: "94", unicode: "128490" },
    { f: "Webdings", code: "95", unicode: "128491" },
    { f: "Webdings", code: "96", unicode: "11156" },
    { f: "Webdings", code: "97", unicode: "10004" },
    { f: "Webdings", code: "98", unicode: "128690" },
    { f: "Webdings", code: "99", unicode: "11036" },
    { f: "Webdings", code: "100", unicode: "128737" },
    { f: "Webdings", code: "101", unicode: "128230" },
    { f: "Webdings", code: "102", unicode: "128753" },
    { f: "Webdings", code: "103", unicode: "11035" },
    { f: "Webdings", code: "104", unicode: "128657" },
    { f: "Webdings", code: "105", unicode: "128712" },
    { f: "Webdings", code: "106", unicode: "128745" },
    { f: "Webdings", code: "107", unicode: "128752" },
    { f: "Webdings", code: "108", unicode: "128968" },
    { f: "Webdings", code: "109", unicode: "128372" },
    { f: "Webdings", code: "110", unicode: "11044" },
    { f: "Webdings", code: "111", unicode: "128741" },
    { f: "Webdings", code: "112", unicode: "128660" },
    { f: "Webdings", code: "113", unicode: "128472" },
    { f: "Webdings", code: "114", unicode: "128473" },
    { f: "Webdings", code: "115", unicode: "10067" },
    { f: "Webdings", code: "116", unicode: "128754" },
    { f: "Webdings", code: "117", unicode: "128647" },
    { f: "Webdings", code: "118", unicode: "128653" },
    { f: "Webdings", code: "119", unicode: "9971" },
    { f: "Webdings", code: "120", unicode: "10680" },
    { f: "Webdings", code: "121", unicode: "8854" },
    { f: "Webdings", code: "122", unicode: "128685" },
    { f: "Webdings", code: "123", unicode: "128494" },
    { f: "Webdings", code: "124", unicode: "9168" },
    { f: "Webdings", code: "125", unicode: "128495" },
    { f: "Webdings", code: "126", unicode: "128498" },
    { f: "Webdings", code: "128", unicode: "128697" },
    { f: "Webdings", code: "129", unicode: "128698" },
    { f: "Webdings", code: "130", unicode: "128713" },
    { f: "Webdings", code: "131", unicode: "128714" },
    { f: "Webdings", code: "132", unicode: "128700" },
    { f: "Webdings", code: "133", unicode: "128125" },
    { f: "Webdings", code: "134", unicode: "127947" },
    { f: "Webdings", code: "135", unicode: "9975" },
    { f: "Webdings", code: "136", unicode: "127938" },
    { f: "Webdings", code: "137", unicode: "127948" },
    { f: "Webdings", code: "138", unicode: "127946" },
    { f: "Webdings", code: "139", unicode: "127940" },
    { f: "Webdings", code: "140", unicode: "127949" },
    { f: "Webdings", code: "141", unicode: "127950" },
    { f: "Webdings", code: "142", unicode: "128664" },
    { f: "Webdings", code: "143", unicode: "128480" },
    { f: "Webdings", code: "144", unicode: "128738" },
    { f: "Webdings", code: "145", unicode: "128176" },
    { f: "Webdings", code: "146", unicode: "127991" },
    { f: "Webdings", code: "147", unicode: "128179" },
    { f: "Webdings", code: "148", unicode: "128106" },
    { f: "Webdings", code: "149", unicode: "128481" },
    { f: "Webdings", code: "150", unicode: "128482" },
    { f: "Webdings", code: "151", unicode: "128483" },
    { f: "Webdings", code: "152", unicode: "10031" },
    { f: "Webdings", code: "153", unicode: "128388" },
    { f: "Webdings", code: "154", unicode: "128389" },
    { f: "Webdings", code: "155", unicode: "128387" },
    { f: "Webdings", code: "156", unicode: "128390" },
    { f: "Webdings", code: "157", unicode: "128441" },
    { f: "Webdings", code: "158", unicode: "128442" },
    { f: "Webdings", code: "159", unicode: "128443" },
    { f: "Webdings", code: "160", unicode: "128373" },
    { f: "Webdings", code: "161", unicode: "128368" },
    { f: "Webdings", code: "162", unicode: "128445" },
    { f: "Webdings", code: "163", unicode: "128446" },
    { f: "Webdings", code: "164", unicode: "128203" },
    { f: "Webdings", code: "165", unicode: "128466" },
    { f: "Webdings", code: "166", unicode: "128467" },
    { f: "Webdings", code: "167", unicode: "128366" },
    { f: "Webdings", code: "168", unicode: "128218" },
    { f: "Webdings", code: "169", unicode: "128478" },
    { f: "Webdings", code: "170", unicode: "128479" },
    { f: "Webdings", code: "171", unicode: "128451" },
    { f: "Webdings", code: "172", unicode: "128450" },
    { f: "Webdings", code: "173", unicode: "128444" },
    { f: "Webdings", code: "174", unicode: "127917" },
    { f: "Webdings", code: "175", unicode: "127900" },
    { f: "Webdings", code: "176", unicode: "127896" },
    { f: "Webdings", code: "177", unicode: "127897" },
    { f: "Webdings", code: "178", unicode: "127911" },
    { f: "Webdings", code: "179", unicode: "128191" },
    { f: "Webdings", code: "180", unicode: "127902" },
    { f: "Webdings", code: "181", unicode: "128247" },
    { f: "Webdings", code: "182", unicode: "127903" },
    { f: "Webdings", code: "183", unicode: "127916" },
    { f: "Webdings", code: "184", unicode: "128253" },
    { f: "Webdings", code: "185", unicode: "128249" },
    { f: "Webdings", code: "186", unicode: "128254" },
    { f: "Webdings", code: "187", unicode: "128251" },
    { f: "Webdings", code: "188", unicode: "127898" },
    { f: "Webdings", code: "189", unicode: "127899" },
    { f: "Webdings", code: "190", unicode: "128250" },
    { f: "Webdings", code: "191", unicode: "128187" },
    { f: "Webdings", code: "192", unicode: "128421" },
    { f: "Webdings", code: "193", unicode: "128422" },
    { f: "Webdings", code: "194", unicode: "128423" },
    { f: "Webdings", code: "195", unicode: "128377" },
    { f: "Webdings", code: "196", unicode: "127918" },
    { f: "Webdings", code: "197", unicode: "128379" },
    { f: "Webdings", code: "198", unicode: "128380" },
    { f: "Webdings", code: "199", unicode: "128223" },
    { f: "Webdings", code: "200", unicode: "128385" },
    { f: "Webdings", code: "201", unicode: "128384" },
    { f: "Webdings", code: "202", unicode: "128424" },
    { f: "Webdings", code: "203", unicode: "128425" },
    { f: "Webdings", code: "204", unicode: "128447" },
    { f: "Webdings", code: "205", unicode: "128426" },
    { f: "Webdings", code: "206", unicode: "128476" },
    { f: "Webdings", code: "207", unicode: "128274" },
    { f: "Webdings", code: "208", unicode: "128275" },
    { f: "Webdings", code: "209", unicode: "128477" },
    { f: "Webdings", code: "210", unicode: "128229" },
    { f: "Webdings", code: "211", unicode: "128228" },
    { f: "Webdings", code: "212", unicode: "128371" },
    { f: "Webdings", code: "213", unicode: "127779" },
    { f: "Webdings", code: "214", unicode: "127780" },
    { f: "Webdings", code: "215", unicode: "127781" },
    { f: "Webdings", code: "216", unicode: "127782" },
    { f: "Webdings", code: "217", unicode: "9729" },
    { f: "Webdings", code: "218", unicode: "127784" },
    { f: "Webdings", code: "219", unicode: "127783" },
    { f: "Webdings", code: "220", unicode: "127785" },
    { f: "Webdings", code: "221", unicode: "127786" },
    { f: "Webdings", code: "222", unicode: "127788" },
    { f: "Webdings", code: "223", unicode: "127787" },
    { f: "Webdings", code: "224", unicode: "127772" },
    { f: "Webdings", code: "225", unicode: "127777" },
    { f: "Webdings", code: "226", unicode: "128715" },
    { f: "Webdings", code: "227", unicode: "128719" },
    { f: "Webdings", code: "228", unicode: "127869" },
    { f: "Webdings", code: "229", unicode: "127864" },
    { f: "Webdings", code: "230", unicode: "128718" },
    { f: "Webdings", code: "231", unicode: "128717" },
    { f: "Webdings", code: "232", unicode: "9413" },
    { f: "Webdings", code: "233", unicode: "9855" },
    { f: "Webdings", code: "234", unicode: "128710" },
    { f: "Webdings", code: "235", unicode: "128392" },
    { f: "Webdings", code: "236", unicode: "127891" },
    { f: "Webdings", code: "237", unicode: "128484" },
    { f: "Webdings", code: "238", unicode: "128485" },
    { f: "Webdings", code: "239", unicode: "128486" },
    { f: "Webdings", code: "240", unicode: "128487" },
    { f: "Webdings", code: "241", unicode: "128746" },
    { f: "Webdings", code: "242", unicode: "128063" },
    { f: "Webdings", code: "243", unicode: "128038" },
    { f: "Webdings", code: "244", unicode: "128031" },
    { f: "Webdings", code: "245", unicode: "128021" },
    { f: "Webdings", code: "246", unicode: "128008" },
    { f: "Webdings", code: "247", unicode: "128620" },
    { f: "Webdings", code: "248", unicode: "128622" },
    { f: "Webdings", code: "249", unicode: "128621" },
    { f: "Webdings", code: "250", unicode: "128623" },
    { f: "Webdings", code: "251", unicode: "128506" },
    { f: "Webdings", code: "252", unicode: "127757" },
    { f: "Webdings", code: "253", unicode: "127759" },
    { f: "Webdings", code: "254", unicode: "127758" },
    { f: "Webdings", code: "255", unicode: "128330" },
    { f: "Wingdings", code: "32", unicode: "32" },
    { f: "Wingdings", code: "33", unicode: "128393" },
    { f: "Wingdings", code: "34", unicode: "9986" },
    { f: "Wingdings", code: "35", unicode: "9985" },
    { f: "Wingdings", code: "36", unicode: "128083" },
    { f: "Wingdings", code: "37", unicode: "128365" },
    { f: "Wingdings", code: "38", unicode: "128366" },
    { f: "Wingdings", code: "39", unicode: "128367" },
    { f: "Wingdings", code: "40", unicode: "128383" },
    { f: "Wingdings", code: "41", unicode: "9990" },
    { f: "Wingdings", code: "42", unicode: "128386" },
    { f: "Wingdings", code: "43", unicode: "128387" },
    { f: "Wingdings", code: "44", unicode: "128234" },
    { f: "Wingdings", code: "45", unicode: "128235" },
    { f: "Wingdings", code: "46", unicode: "128236" },
    { f: "Wingdings", code: "47", unicode: "128237" },
    { f: "Wingdings", code: "48", unicode: "128448" },
    { f: "Wingdings", code: "49", unicode: "128449" },
    { f: "Wingdings", code: "50", unicode: "128462" },
    { f: "Wingdings", code: "51", unicode: "128463" },
    { f: "Wingdings", code: "52", unicode: "128464" },
    { f: "Wingdings", code: "53", unicode: "128452" },
    { f: "Wingdings", code: "54", unicode: "8987" },
    { f: "Wingdings", code: "55", unicode: "128430" },
    { f: "Wingdings", code: "56", unicode: "128432" },
    { f: "Wingdings", code: "57", unicode: "128434" },
    { f: "Wingdings", code: "58", unicode: "128435" },
    { f: "Wingdings", code: "59", unicode: "128436" },
    { f: "Wingdings", code: "60", unicode: "128427" },
    { f: "Wingdings", code: "61", unicode: "128428" },
    { f: "Wingdings", code: "62", unicode: "9991" },
    { f: "Wingdings", code: "63", unicode: "9997" },
    { f: "Wingdings", code: "64", unicode: "128398" },
    { f: "Wingdings", code: "65", unicode: "9996" },
    { f: "Wingdings", code: "66", unicode: "128399" },
    { f: "Wingdings", code: "67", unicode: "128077" },
    { f: "Wingdings", code: "68", unicode: "128078" },
    { f: "Wingdings", code: "69", unicode: "9756" },
    { f: "Wingdings", code: "70", unicode: "9758" },
    { f: "Wingdings", code: "71", unicode: "9757" },
    { f: "Wingdings", code: "72", unicode: "9759" },
    { f: "Wingdings", code: "73", unicode: "128400" },
    { f: "Wingdings", code: "74", unicode: "9786" },
    { f: "Wingdings", code: "75", unicode: "128528" },
    { f: "Wingdings", code: "76", unicode: "9785" },
    { f: "Wingdings", code: "77", unicode: "128163" },
    { f: "Wingdings", code: "78", unicode: "128369" },
    { f: "Wingdings", code: "79", unicode: "127987" },
    { f: "Wingdings", code: "80", unicode: "127985" },
    { f: "Wingdings", code: "81", unicode: "9992" },
    { f: "Wingdings", code: "82", unicode: "9788" },
    { f: "Wingdings", code: "83", unicode: "127778" },
    { f: "Wingdings", code: "84", unicode: "10052" },
    { f: "Wingdings", code: "85", unicode: "128326" },
    { f: "Wingdings", code: "86", unicode: "10014" },
    { f: "Wingdings", code: "87", unicode: "128328" },
    { f: "Wingdings", code: "88", unicode: "10016" },
    { f: "Wingdings", code: "89", unicode: "10017" },
    { f: "Wingdings", code: "90", unicode: "9770" },
    { f: "Wingdings", code: "91", unicode: "9775" },
    { f: "Wingdings", code: "92", unicode: "128329" },
    { f: "Wingdings", code: "93", unicode: "9784" },
    { f: "Wingdings", code: "94", unicode: "9800" },
    { f: "Wingdings", code: "95", unicode: "9801" },
    { f: "Wingdings", code: "96", unicode: "9802" },
    { f: "Wingdings", code: "97", unicode: "9803" },
    { f: "Wingdings", code: "98", unicode: "9804" },
    { f: "Wingdings", code: "99", unicode: "9805" },
    { f: "Wingdings", code: "100", unicode: "9806" },
    { f: "Wingdings", code: "101", unicode: "9807" },
    { f: "Wingdings", code: "102", unicode: "9808" },
    { f: "Wingdings", code: "103", unicode: "9809" },
    { f: "Wingdings", code: "104", unicode: "9810" },
    { f: "Wingdings", code: "105", unicode: "9811" },
    { f: "Wingdings", code: "106", unicode: "128624" },
    { f: "Wingdings", code: "107", unicode: "128629" },
    { f: "Wingdings", code: "108", unicode: "9899" },
    { f: "Wingdings", code: "109", unicode: "128318" },
    { f: "Wingdings", code: "110", unicode: "9724" },
    { f: "Wingdings", code: "111", unicode: "128911" },
    { f: "Wingdings", code: "112", unicode: "128912" },
    { f: "Wingdings", code: "113", unicode: "10065" },
    { f: "Wingdings", code: "114", unicode: "10066" },
    { f: "Wingdings", code: "115", unicode: "128927" },
    { f: "Wingdings", code: "116", unicode: "10731" },
    { f: "Wingdings", code: "117", unicode: "9670" },
    { f: "Wingdings", code: "118", unicode: "10070" },
    { f: "Wingdings", code: "119", unicode: "11049" },
    { f: "Wingdings", code: "120", unicode: "8999" },
    { f: "Wingdings", code: "121", unicode: "11193" },
    { f: "Wingdings", code: "122", unicode: "8984" },
    { f: "Wingdings", code: "123", unicode: "127989" },
    { f: "Wingdings", code: "124", unicode: "127990" },
    { f: "Wingdings", code: "125", unicode: "128630" },
    { f: "Wingdings", code: "126", unicode: "128631" },
    { f: "Wingdings", code: "127", unicode: "9647" },
    { f: "Wingdings", code: "128", unicode: "127243" },
    { f: "Wingdings", code: "129", unicode: "10112" },
    { f: "Wingdings", code: "130", unicode: "10113" },
    { f: "Wingdings", code: "131", unicode: "10114" },
    { f: "Wingdings", code: "132", unicode: "10115" },
    { f: "Wingdings", code: "133", unicode: "10116" },
    { f: "Wingdings", code: "134", unicode: "10117" },
    { f: "Wingdings", code: "135", unicode: "10118" },
    { f: "Wingdings", code: "136", unicode: "10119" },
    { f: "Wingdings", code: "137", unicode: "10120" },
    { f: "Wingdings", code: "138", unicode: "10121" },
    { f: "Wingdings", code: "139", unicode: "127244" },
    { f: "Wingdings", code: "140", unicode: "10122" },
    { f: "Wingdings", code: "141", unicode: "10123" },
    { f: "Wingdings", code: "142", unicode: "10124" },
    { f: "Wingdings", code: "143", unicode: "10125" },
    { f: "Wingdings", code: "144", unicode: "10126" },
    { f: "Wingdings", code: "145", unicode: "10127" },
    { f: "Wingdings", code: "146", unicode: "10128" },
    { f: "Wingdings", code: "147", unicode: "10129" },
    { f: "Wingdings", code: "148", unicode: "10130" },
    { f: "Wingdings", code: "149", unicode: "10131" },
    { f: "Wingdings", code: "150", unicode: "128610" },
    { f: "Wingdings", code: "151", unicode: "128608" },
    { f: "Wingdings", code: "152", unicode: "128609" },
    { f: "Wingdings", code: "153", unicode: "128611" },
    { f: "Wingdings", code: "154", unicode: "128606" },
    { f: "Wingdings", code: "155", unicode: "128604" },
    { f: "Wingdings", code: "156", unicode: "128605" },
    { f: "Wingdings", code: "157", unicode: "128607" },
    { f: "Wingdings", code: "158", unicode: "8729" },
    { f: "Wingdings", code: "159", unicode: "8226" },
    { f: "Wingdings", code: "160", unicode: "11037" },
    { f: "Wingdings", code: "161", unicode: "11096" },
    { f: "Wingdings", code: "162", unicode: "128902" },
    { f: "Wingdings", code: "163", unicode: "128904" },
    { f: "Wingdings", code: "164", unicode: "128906" },
    { f: "Wingdings", code: "165", unicode: "128907" },
    { f: "Wingdings", code: "166", unicode: "128319" },
    { f: "Wingdings", code: "167", unicode: "9642" },
    { f: "Wingdings", code: "168", unicode: "128910" },
    { f: "Wingdings", code: "169", unicode: "128961" },
    { f: "Wingdings", code: "170", unicode: "128965" },
    { f: "Wingdings", code: "171", unicode: "9733" },
    { f: "Wingdings", code: "172", unicode: "128971" },
    { f: "Wingdings", code: "173", unicode: "128975" },
    { f: "Wingdings", code: "174", unicode: "128979" },
    { f: "Wingdings", code: "175", unicode: "128977" },
    { f: "Wingdings", code: "176", unicode: "11216" },
    { f: "Wingdings", code: "177", unicode: "8982" },
    { f: "Wingdings", code: "178", unicode: "11214" },
    { f: "Wingdings", code: "179", unicode: "11215" },
    { f: "Wingdings", code: "180", unicode: "11217" },
    { f: "Wingdings", code: "181", unicode: "10026" },
    { f: "Wingdings", code: "182", unicode: "10032" },
    { f: "Wingdings", code: "183", unicode: "128336" },
    { f: "Wingdings", code: "184", unicode: "128337" },
    { f: "Wingdings", code: "185", unicode: "128338" },
    { f: "Wingdings", code: "186", unicode: "128339" },
    { f: "Wingdings", code: "187", unicode: "128340" },
    { f: "Wingdings", code: "188", unicode: "128341" },
    { f: "Wingdings", code: "189", unicode: "128342" },
    { f: "Wingdings", code: "190", unicode: "128343" },
    { f: "Wingdings", code: "191", unicode: "128344" },
    { f: "Wingdings", code: "192", unicode: "128345" },
    { f: "Wingdings", code: "193", unicode: "128346" },
    { f: "Wingdings", code: "194", unicode: "128347" },
    { f: "Wingdings", code: "195", unicode: "11184" },
    { f: "Wingdings", code: "196", unicode: "11185" },
    { f: "Wingdings", code: "197", unicode: "11186" },
    { f: "Wingdings", code: "198", unicode: "11187" },
    { f: "Wingdings", code: "199", unicode: "11188" },
    { f: "Wingdings", code: "200", unicode: "11189" },
    { f: "Wingdings", code: "201", unicode: "11190" },
    { f: "Wingdings", code: "202", unicode: "11191" },
    { f: "Wingdings", code: "203", unicode: "128618" },
    { f: "Wingdings", code: "204", unicode: "128619" },
    { f: "Wingdings", code: "205", unicode: "128597" },
    { f: "Wingdings", code: "206", unicode: "128596" },
    { f: "Wingdings", code: "207", unicode: "128599" },
    { f: "Wingdings", code: "208", unicode: "128598" },
    { f: "Wingdings", code: "209", unicode: "128592" },
    { f: "Wingdings", code: "210", unicode: "128593" },
    { f: "Wingdings", code: "211", unicode: "128594" },
    { f: "Wingdings", code: "212", unicode: "128595" },
    { f: "Wingdings", code: "213", unicode: "9003" },
    { f: "Wingdings", code: "214", unicode: "8998" },
    { f: "Wingdings", code: "215", unicode: "11160" },
    { f: "Wingdings", code: "216", unicode: "11162" },
    { f: "Wingdings", code: "217", unicode: "11161" },
    { f: "Wingdings", code: "218", unicode: "11163" },
    { f: "Wingdings", code: "219", unicode: "11144" },
    { f: "Wingdings", code: "220", unicode: "11146" },
    { f: "Wingdings", code: "221", unicode: "11145" },
    { f: "Wingdings", code: "222", unicode: "11147" },
    { f: "Wingdings", code: "223", unicode: "129128" },
    { f: "Wingdings", code: "224", unicode: "129130" },
    { f: "Wingdings", code: "225", unicode: "129129" },
    { f: "Wingdings", code: "226", unicode: "129131" },
    { f: "Wingdings", code: "227", unicode: "129132" },
    { f: "Wingdings", code: "228", unicode: "129133" },
    { f: "Wingdings", code: "229", unicode: "129135" },
    { f: "Wingdings", code: "230", unicode: "129134" },
    { f: "Wingdings", code: "231", unicode: "129144" },
    { f: "Wingdings", code: "232", unicode: "129146" },
    { f: "Wingdings", code: "233", unicode: "129145" },
    { f: "Wingdings", code: "234", unicode: "129147" },
    { f: "Wingdings", code: "235", unicode: "129148" },
    { f: "Wingdings", code: "236", unicode: "129149" },
    { f: "Wingdings", code: "237", unicode: "129151" },
    { f: "Wingdings", code: "238", unicode: "129150" },
    { f: "Wingdings", code: "239", unicode: "8678" },
    { f: "Wingdings", code: "240", unicode: "8680" },
    { f: "Wingdings", code: "241", unicode: "8679" },
    { f: "Wingdings", code: "242", unicode: "8681" },
    { f: "Wingdings", code: "243", unicode: "11012" },
    { f: "Wingdings", code: "244", unicode: "8691" },
    { f: "Wingdings", code: "245", unicode: "11009" },
    { f: "Wingdings", code: "246", unicode: "11008" },
    { f: "Wingdings", code: "247", unicode: "11011" },
    { f: "Wingdings", code: "248", unicode: "11010" },
    { f: "Wingdings", code: "249", unicode: "129196" },
    { f: "Wingdings", code: "250", unicode: "129197" },
    { f: "Wingdings", code: "251", unicode: "128502" },
    { f: "Wingdings", code: "252", unicode: "10003" },
    { f: "Wingdings", code: "253", unicode: "128503" },
    { f: "Wingdings", code: "254", unicode: "128505" },
    { f: "Wingdings 2", code: "32", unicode: "32" },
    { f: "Wingdings 2", code: "33", unicode: "128394" },
    { f: "Wingdings 2", code: "34", unicode: "128395" },
    { f: "Wingdings 2", code: "35", unicode: "128396" },
    { f: "Wingdings 2", code: "36", unicode: "128397" },
    { f: "Wingdings 2", code: "37", unicode: "9988" },
    { f: "Wingdings 2", code: "38", unicode: "9984" },
    { f: "Wingdings 2", code: "39", unicode: "128382" },
    { f: "Wingdings 2", code: "40", unicode: "128381" },
    { f: "Wingdings 2", code: "41", unicode: "128453" },
    { f: "Wingdings 2", code: "42", unicode: "128454" },
    { f: "Wingdings 2", code: "43", unicode: "128455" },
    { f: "Wingdings 2", code: "44", unicode: "128456" },
    { f: "Wingdings 2", code: "45", unicode: "128457" },
    { f: "Wingdings 2", code: "46", unicode: "128458" },
    { f: "Wingdings 2", code: "47", unicode: "128459" },
    { f: "Wingdings 2", code: "48", unicode: "128460" },
    { f: "Wingdings 2", code: "49", unicode: "128461" },
    { f: "Wingdings 2", code: "50", unicode: "128203" },
    { f: "Wingdings 2", code: "51", unicode: "128465" },
    { f: "Wingdings 2", code: "52", unicode: "128468" },
    { f: "Wingdings 2", code: "53", unicode: "128437" },
    { f: "Wingdings 2", code: "54", unicode: "128438" },
    { f: "Wingdings 2", code: "55", unicode: "128439" },
    { f: "Wingdings 2", code: "56", unicode: "128440" },
    { f: "Wingdings 2", code: "57", unicode: "128429" },
    { f: "Wingdings 2", code: "58", unicode: "128431" },
    { f: "Wingdings 2", code: "59", unicode: "128433" },
    { f: "Wingdings 2", code: "60", unicode: "128402" },
    { f: "Wingdings 2", code: "61", unicode: "128403" },
    { f: "Wingdings 2", code: "62", unicode: "128408" },
    { f: "Wingdings 2", code: "63", unicode: "128409" },
    { f: "Wingdings 2", code: "64", unicode: "128410" },
    { f: "Wingdings 2", code: "65", unicode: "128411" },
    { f: "Wingdings 2", code: "66", unicode: "128072" },
    { f: "Wingdings 2", code: "67", unicode: "128073" },
    { f: "Wingdings 2", code: "68", unicode: "128412" },
    { f: "Wingdings 2", code: "69", unicode: "128413" },
    { f: "Wingdings 2", code: "70", unicode: "128414" },
    { f: "Wingdings 2", code: "71", unicode: "128415" },
    { f: "Wingdings 2", code: "72", unicode: "128416" },
    { f: "Wingdings 2", code: "73", unicode: "128417" },
    { f: "Wingdings 2", code: "74", unicode: "128070" },
    { f: "Wingdings 2", code: "75", unicode: "128071" },
    { f: "Wingdings 2", code: "76", unicode: "128418" },
    { f: "Wingdings 2", code: "77", unicode: "128419" },
    { f: "Wingdings 2", code: "78", unicode: "128401" },
    { f: "Wingdings 2", code: "79", unicode: "128500" },
    { f: "Wingdings 2", code: "80", unicode: "128504" },
    { f: "Wingdings 2", code: "81", unicode: "128501" },
    { f: "Wingdings 2", code: "82", unicode: "9745" },
    { f: "Wingdings 2", code: "83", unicode: "11197" },
    { f: "Wingdings 2", code: "84", unicode: "9746" },
    { f: "Wingdings 2", code: "85", unicode: "11198" },
    { f: "Wingdings 2", code: "86", unicode: "11199" },
    { f: "Wingdings 2", code: "87", unicode: "128711" },
    { f: "Wingdings 2", code: "88", unicode: "10680" },
    { f: "Wingdings 2", code: "89", unicode: "128625" },
    { f: "Wingdings 2", code: "90", unicode: "128628" },
    { f: "Wingdings 2", code: "91", unicode: "128626" },
    { f: "Wingdings 2", code: "92", unicode: "128627" },
    { f: "Wingdings 2", code: "93", unicode: "8253" },
    { f: "Wingdings 2", code: "94", unicode: "128633" },
    { f: "Wingdings 2", code: "95", unicode: "128634" },
    { f: "Wingdings 2", code: "96", unicode: "128635" },
    { f: "Wingdings 2", code: "97", unicode: "128614" },
    { f: "Wingdings 2", code: "98", unicode: "128612" },
    { f: "Wingdings 2", code: "99", unicode: "128613" },
    { f: "Wingdings 2", code: "100", unicode: "128615" },
    { f: "Wingdings 2", code: "101", unicode: "128602" },
    { f: "Wingdings 2", code: "102", unicode: "128600" },
    { f: "Wingdings 2", code: "103", unicode: "128601" },
    { f: "Wingdings 2", code: "104", unicode: "128603" },
    { f: "Wingdings 2", code: "105", unicode: "9450" },
    { f: "Wingdings 2", code: "106", unicode: "9312" },
    { f: "Wingdings 2", code: "107", unicode: "9313" },
    { f: "Wingdings 2", code: "108", unicode: "9314" },
    { f: "Wingdings 2", code: "109", unicode: "9315" },
    { f: "Wingdings 2", code: "110", unicode: "9316" },
    { f: "Wingdings 2", code: "111", unicode: "9317" },
    { f: "Wingdings 2", code: "112", unicode: "9318" },
    { f: "Wingdings 2", code: "113", unicode: "9319" },
    { f: "Wingdings 2", code: "114", unicode: "9320" },
    { f: "Wingdings 2", code: "115", unicode: "9321" },
    { f: "Wingdings 2", code: "116", unicode: "9471" },
    { f: "Wingdings 2", code: "117", unicode: "10102" },
    { f: "Wingdings 2", code: "118", unicode: "10103" },
    { f: "Wingdings 2", code: "119", unicode: "10104" },
    { f: "Wingdings 2", code: "120", unicode: "10105" },
    { f: "Wingdings 2", code: "121", unicode: "10106" },
    { f: "Wingdings 2", code: "122", unicode: "10107" },
    { f: "Wingdings 2", code: "123", unicode: "10108" },
    { f: "Wingdings 2", code: "124", unicode: "10109" },
    { f: "Wingdings 2", code: "125", unicode: "10110" },
    { f: "Wingdings 2", code: "126", unicode: "10111" },
    { f: "Wingdings 2", code: "128", unicode: "9737" },
    { f: "Wingdings 2", code: "129", unicode: "127765" },
    { f: "Wingdings 2", code: "130", unicode: "9789" },
    { f: "Wingdings 2", code: "131", unicode: "9790" },
    { f: "Wingdings 2", code: "132", unicode: "11839" },
    { f: "Wingdings 2", code: "133", unicode: "10013" },
    { f: "Wingdings 2", code: "134", unicode: "128327" },
    { f: "Wingdings 2", code: "135", unicode: "128348" },
    { f: "Wingdings 2", code: "136", unicode: "128349" },
    { f: "Wingdings 2", code: "137", unicode: "128350" },
    { f: "Wingdings 2", code: "138", unicode: "128351" },
    { f: "Wingdings 2", code: "139", unicode: "128352" },
    { f: "Wingdings 2", code: "140", unicode: "128353" },
    { f: "Wingdings 2", code: "141", unicode: "128354" },
    { f: "Wingdings 2", code: "142", unicode: "128355" },
    { f: "Wingdings 2", code: "143", unicode: "128356" },
    { f: "Wingdings 2", code: "144", unicode: "128357" },
    { f: "Wingdings 2", code: "145", unicode: "128358" },
    { f: "Wingdings 2", code: "146", unicode: "128359" },
    { f: "Wingdings 2", code: "147", unicode: "128616" },
    { f: "Wingdings 2", code: "148", unicode: "128617" },
    { f: "Wingdings 2", code: "149", unicode: "8901" },
    { f: "Wingdings 2", code: "150", unicode: "128900" },
    { f: "Wingdings 2", code: "151", unicode: "10625" },
    { f: "Wingdings 2", code: "152", unicode: "9679" },
    { f: "Wingdings 2", code: "153", unicode: "9675" },
    { f: "Wingdings 2", code: "154", unicode: "128901" },
    { f: "Wingdings 2", code: "155", unicode: "128903" },
    { f: "Wingdings 2", code: "156", unicode: "128905" },
    { f: "Wingdings 2", code: "157", unicode: "8857" },
    { f: "Wingdings 2", code: "158", unicode: "10687" },
    { f: "Wingdings 2", code: "159", unicode: "128908" },
    { f: "Wingdings 2", code: "160", unicode: "128909" },
    { f: "Wingdings 2", code: "161", unicode: "9726" },
    { f: "Wingdings 2", code: "162", unicode: "9632" },
    { f: "Wingdings 2", code: "163", unicode: "9633" },
    { f: "Wingdings 2", code: "164", unicode: "128913" },
    { f: "Wingdings 2", code: "165", unicode: "128914" },
    { f: "Wingdings 2", code: "166", unicode: "128915" },
    { f: "Wingdings 2", code: "167", unicode: "128916" },
    { f: "Wingdings 2", code: "168", unicode: "9635" },
    { f: "Wingdings 2", code: "169", unicode: "128917" },
    { f: "Wingdings 2", code: "170", unicode: "128918" },
    { f: "Wingdings 2", code: "171", unicode: "128919" },
    { f: "Wingdings 2", code: "172", unicode: "128920" },
    { f: "Wingdings 2", code: "173", unicode: "11049" },
    { f: "Wingdings 2", code: "174", unicode: "11045" },
    { f: "Wingdings 2", code: "175", unicode: "9671" },
    { f: "Wingdings 2", code: "176", unicode: "128922" },
    { f: "Wingdings 2", code: "177", unicode: "9672" },
    { f: "Wingdings 2", code: "178", unicode: "128923" },
    { f: "Wingdings 2", code: "179", unicode: "128924" },
    { f: "Wingdings 2", code: "180", unicode: "128925" },
    { f: "Wingdings 2", code: "181", unicode: "128926" },
    { f: "Wingdings 2", code: "182", unicode: "11050" },
    { f: "Wingdings 2", code: "183", unicode: "11047" },
    { f: "Wingdings 2", code: "184", unicode: "9674" },
    { f: "Wingdings 2", code: "185", unicode: "128928" },
    { f: "Wingdings 2", code: "186", unicode: "9686" },
    { f: "Wingdings 2", code: "187", unicode: "9687" },
    { f: "Wingdings 2", code: "188", unicode: "11210" },
    { f: "Wingdings 2", code: "189", unicode: "11211" },
    { f: "Wingdings 2", code: "190", unicode: "11200" },
    { f: "Wingdings 2", code: "191", unicode: "11201" },
    { f: "Wingdings 2", code: "192", unicode: "11039" },
    { f: "Wingdings 2", code: "193", unicode: "11202" },
    { f: "Wingdings 2", code: "194", unicode: "11043" },
    { f: "Wingdings 2", code: "195", unicode: "11042" },
    { f: "Wingdings 2", code: "196", unicode: "11203" },
    { f: "Wingdings 2", code: "197", unicode: "11204" },
    { f: "Wingdings 2", code: "198", unicode: "128929" },
    { f: "Wingdings 2", code: "199", unicode: "128930" },
    { f: "Wingdings 2", code: "200", unicode: "128931" },
    { f: "Wingdings 2", code: "201", unicode: "128932" },
    { f: "Wingdings 2", code: "202", unicode: "128933" },
    { f: "Wingdings 2", code: "203", unicode: "128934" },
    { f: "Wingdings 2", code: "204", unicode: "128935" },
    { f: "Wingdings 2", code: "205", unicode: "128936" },
    { f: "Wingdings 2", code: "206", unicode: "128937" },
    { f: "Wingdings 2", code: "207", unicode: "128938" },
    { f: "Wingdings 2", code: "208", unicode: "128939" },
    { f: "Wingdings 2", code: "209", unicode: "128940" },
    { f: "Wingdings 2", code: "210", unicode: "128941" },
    { f: "Wingdings 2", code: "211", unicode: "128942" },
    { f: "Wingdings 2", code: "212", unicode: "128943" },
    { f: "Wingdings 2", code: "213", unicode: "128944" },
    { f: "Wingdings 2", code: "214", unicode: "128945" },
    { f: "Wingdings 2", code: "215", unicode: "128946" },
    { f: "Wingdings 2", code: "216", unicode: "128947" },
    { f: "Wingdings 2", code: "217", unicode: "128948" },
    { f: "Wingdings 2", code: "218", unicode: "128949" },
    { f: "Wingdings 2", code: "219", unicode: "128950" },
    { f: "Wingdings 2", code: "220", unicode: "128951" },
    { f: "Wingdings 2", code: "221", unicode: "128952" },
    { f: "Wingdings 2", code: "222", unicode: "128953" },
    { f: "Wingdings 2", code: "223", unicode: "128954" },
    { f: "Wingdings 2", code: "224", unicode: "128955" },
    { f: "Wingdings 2", code: "225", unicode: "128956" },
    { f: "Wingdings 2", code: "226", unicode: "128957" },
    { f: "Wingdings 2", code: "227", unicode: "128958" },
    { f: "Wingdings 2", code: "228", unicode: "128959" },
    { f: "Wingdings 2", code: "229", unicode: "128960" },
    { f: "Wingdings 2", code: "230", unicode: "128962" },
    { f: "Wingdings 2", code: "231", unicode: "128964" },
    { f: "Wingdings 2", code: "232", unicode: "128966" },
    { f: "Wingdings 2", code: "233", unicode: "128969" },
    { f: "Wingdings 2", code: "234", unicode: "128970" },
    { f: "Wingdings 2", code: "235", unicode: "10038" },
    { f: "Wingdings 2", code: "236", unicode: "128972" },
    { f: "Wingdings 2", code: "237", unicode: "128974" },
    { f: "Wingdings 2", code: "238", unicode: "128976" },
    { f: "Wingdings 2", code: "239", unicode: "128978" },
    { f: "Wingdings 2", code: "240", unicode: "10041" },
    { f: "Wingdings 2", code: "241", unicode: "128963" },
    { f: "Wingdings 2", code: "242", unicode: "128967" },
    { f: "Wingdings 2", code: "243", unicode: "10031" },
    { f: "Wingdings 2", code: "244", unicode: "128973" },
    { f: "Wingdings 2", code: "245", unicode: "128980" },
    { f: "Wingdings 2", code: "246", unicode: "11212" },
    { f: "Wingdings 2", code: "247", unicode: "11213" },
    { f: "Wingdings 2", code: "248", unicode: "8258" },
    { f: "Wingdings 3", code: "32", unicode: "32" },
    { f: "Wingdings 3", code: "33", unicode: "11104" },
    { f: "Wingdings 3", code: "34", unicode: "11106" },
    { f: "Wingdings 3", code: "35", unicode: "11105" },
    { f: "Wingdings 3", code: "36", unicode: "11107" },
    { f: "Wingdings 3", code: "37", unicode: "11110" },
    { f: "Wingdings 3", code: "38", unicode: "11111" },
    { f: "Wingdings 3", code: "39", unicode: "11113" },
    { f: "Wingdings 3", code: "40", unicode: "11112" },
    { f: "Wingdings 3", code: "41", unicode: "11120" },
    { f: "Wingdings 3", code: "42", unicode: "11122" },
    { f: "Wingdings 3", code: "43", unicode: "11121" },
    { f: "Wingdings 3", code: "44", unicode: "11123" },
    { f: "Wingdings 3", code: "45", unicode: "11126" },
    { f: "Wingdings 3", code: "46", unicode: "11128" },
    { f: "Wingdings 3", code: "47", unicode: "11131" },
    { f: "Wingdings 3", code: "48", unicode: "11133" },
    { f: "Wingdings 3", code: "49", unicode: "11108" },
    { f: "Wingdings 3", code: "50", unicode: "11109" },
    { f: "Wingdings 3", code: "51", unicode: "11114" },
    { f: "Wingdings 3", code: "52", unicode: "11116" },
    { f: "Wingdings 3", code: "53", unicode: "11115" },
    { f: "Wingdings 3", code: "54", unicode: "11117" },
    { f: "Wingdings 3", code: "55", unicode: "11085" },
    { f: "Wingdings 3", code: "56", unicode: "11168" },
    { f: "Wingdings 3", code: "57", unicode: "11169" },
    { f: "Wingdings 3", code: "58", unicode: "11170" },
    { f: "Wingdings 3", code: "59", unicode: "11171" },
    { f: "Wingdings 3", code: "60", unicode: "11172" },
    { f: "Wingdings 3", code: "61", unicode: "11173" },
    { f: "Wingdings 3", code: "62", unicode: "11174" },
    { f: "Wingdings 3", code: "63", unicode: "11175" },
    { f: "Wingdings 3", code: "64", unicode: "11152" },
    { f: "Wingdings 3", code: "65", unicode: "11153" },
    { f: "Wingdings 3", code: "66", unicode: "11154" },
    { f: "Wingdings 3", code: "67", unicode: "11155" },
    { f: "Wingdings 3", code: "68", unicode: "11136" },
    { f: "Wingdings 3", code: "69", unicode: "11139" },
    { f: "Wingdings 3", code: "70", unicode: "11134" },
    { f: "Wingdings 3", code: "71", unicode: "11135" },
    { f: "Wingdings 3", code: "72", unicode: "11140" },
    { f: "Wingdings 3", code: "73", unicode: "11142" },
    { f: "Wingdings 3", code: "74", unicode: "11141" },
    { f: "Wingdings 3", code: "75", unicode: "11143" },
    { f: "Wingdings 3", code: "76", unicode: "11151" },
    { f: "Wingdings 3", code: "77", unicode: "11149" },
    { f: "Wingdings 3", code: "78", unicode: "11150" },
    { f: "Wingdings 3", code: "79", unicode: "11148" },
    { f: "Wingdings 3", code: "80", unicode: "11118" },
    { f: "Wingdings 3", code: "81", unicode: "11119" },
    { f: "Wingdings 3", code: "82", unicode: "9099" },
    { f: "Wingdings 3", code: "83", unicode: "8996" },
    { f: "Wingdings 3", code: "84", unicode: "8963" },
    { f: "Wingdings 3", code: "85", unicode: "8997" },
    { f: "Wingdings 3", code: "86", unicode: "9251" },
    { f: "Wingdings 3", code: "87", unicode: "9085" },
    { f: "Wingdings 3", code: "88", unicode: "8682" },
    { f: "Wingdings 3", code: "89", unicode: "11192" },
    { f: "Wingdings 3", code: "90", unicode: "129184" },
    { f: "Wingdings 3", code: "91", unicode: "129185" },
    { f: "Wingdings 3", code: "92", unicode: "129186" },
    { f: "Wingdings 3", code: "93", unicode: "129187" },
    { f: "Wingdings 3", code: "94", unicode: "129188" },
    { f: "Wingdings 3", code: "95", unicode: "129189" },
    { f: "Wingdings 3", code: "96", unicode: "129190" },
    { f: "Wingdings 3", code: "97", unicode: "129191" },
    { f: "Wingdings 3", code: "98", unicode: "129192" },
    { f: "Wingdings 3", code: "99", unicode: "129193" },
    { f: "Wingdings 3", code: "100", unicode: "129194" },
    { f: "Wingdings 3", code: "101", unicode: "129195" },
    { f: "Wingdings 3", code: "102", unicode: "129104" },
    { f: "Wingdings 3", code: "103", unicode: "129106" },
    { f: "Wingdings 3", code: "104", unicode: "129105" },
    { f: "Wingdings 3", code: "105", unicode: "129107" },
    { f: "Wingdings 3", code: "106", unicode: "129108" },
    { f: "Wingdings 3", code: "107", unicode: "129109" },
    { f: "Wingdings 3", code: "108", unicode: "129111" },
    { f: "Wingdings 3", code: "109", unicode: "129110" },
    { f: "Wingdings 3", code: "110", unicode: "129112" },
    { f: "Wingdings 3", code: "111", unicode: "129113" },
    { f: "Wingdings 3", code: "112", unicode: "9650" },
    { f: "Wingdings 3", code: "113", unicode: "9660" },
    { f: "Wingdings 3", code: "114", unicode: "9651" },
    { f: "Wingdings 3", code: "115", unicode: "9661" },
    { f: "Wingdings 3", code: "116", unicode: "9664" },
    { f: "Wingdings 3", code: "117", unicode: "9654" },
    { f: "Wingdings 3", code: "118", unicode: "9665" },
    { f: "Wingdings 3", code: "119", unicode: "9655" },
    { f: "Wingdings 3", code: "120", unicode: "9699" },
    { f: "Wingdings 3", code: "121", unicode: "9698" },
    { f: "Wingdings 3", code: "122", unicode: "9700" },
    { f: "Wingdings 3", code: "123", unicode: "9701" },
    { f: "Wingdings 3", code: "124", unicode: "128896" },
    { f: "Wingdings 3", code: "125", unicode: "128898" },
    { f: "Wingdings 3", code: "126", unicode: "128897" },
    { f: "Wingdings 3", code: "128", unicode: "128899" },
    { f: "Wingdings 3", code: "129", unicode: "11205" },
    { f: "Wingdings 3", code: "130", unicode: "11206" },
    { f: "Wingdings 3", code: "131", unicode: "11207" },
    { f: "Wingdings 3", code: "132", unicode: "11208" },
    { f: "Wingdings 3", code: "133", unicode: "11164" },
    { f: "Wingdings 3", code: "134", unicode: "11166" },
    { f: "Wingdings 3", code: "135", unicode: "11165" },
    { f: "Wingdings 3", code: "136", unicode: "11167" },
    { f: "Wingdings 3", code: "137", unicode: "129040" },
    { f: "Wingdings 3", code: "138", unicode: "129042" },
    { f: "Wingdings 3", code: "139", unicode: "129041" },
    { f: "Wingdings 3", code: "140", unicode: "129043" },
    { f: "Wingdings 3", code: "141", unicode: "129044" },
    { f: "Wingdings 3", code: "142", unicode: "129046" },
    { f: "Wingdings 3", code: "143", unicode: "129045" },
    { f: "Wingdings 3", code: "144", unicode: "129047" },
    { f: "Wingdings 3", code: "145", unicode: "129048" },
    { f: "Wingdings 3", code: "146", unicode: "129050" },
    { f: "Wingdings 3", code: "147", unicode: "129049" },
    { f: "Wingdings 3", code: "148", unicode: "129051" },
    { f: "Wingdings 3", code: "149", unicode: "129052" },
    { f: "Wingdings 3", code: "150", unicode: "129054" },
    { f: "Wingdings 3", code: "151", unicode: "129053" },
    { f: "Wingdings 3", code: "152", unicode: "129055" },
    { f: "Wingdings 3", code: "153", unicode: "129024" },
    { f: "Wingdings 3", code: "154", unicode: "129026" },
    { f: "Wingdings 3", code: "155", unicode: "129025" },
    { f: "Wingdings 3", code: "156", unicode: "129027" },
    { f: "Wingdings 3", code: "157", unicode: "129028" },
    { f: "Wingdings 3", code: "158", unicode: "129030" },
    { f: "Wingdings 3", code: "159", unicode: "129029" },
    { f: "Wingdings 3", code: "160", unicode: "129031" },
    { f: "Wingdings 3", code: "161", unicode: "129032" },
    { f: "Wingdings 3", code: "162", unicode: "129034" },
    { f: "Wingdings 3", code: "163", unicode: "129033" },
    { f: "Wingdings 3", code: "164", unicode: "129035" },
    { f: "Wingdings 3", code: "165", unicode: "129056" },
    { f: "Wingdings 3", code: "166", unicode: "129058" },
    { f: "Wingdings 3", code: "167", unicode: "129060" },
    { f: "Wingdings 3", code: "168", unicode: "129062" },
    { f: "Wingdings 3", code: "169", unicode: "129064" },
    { f: "Wingdings 3", code: "170", unicode: "129066" },
    { f: "Wingdings 3", code: "171", unicode: "129068" },
    { f: "Wingdings 3", code: "172", unicode: "129180" },
    { f: "Wingdings 3", code: "173", unicode: "129181" },
    { f: "Wingdings 3", code: "174", unicode: "129182" },
    { f: "Wingdings 3", code: "175", unicode: "129183" },
    { f: "Wingdings 3", code: "176", unicode: "129070" },
    { f: "Wingdings 3", code: "177", unicode: "129072" },
    { f: "Wingdings 3", code: "178", unicode: "129074" },
    { f: "Wingdings 3", code: "179", unicode: "129076" },
    { f: "Wingdings 3", code: "180", unicode: "129078" },
    { f: "Wingdings 3", code: "181", unicode: "129080" },
    { f: "Wingdings 3", code: "182", unicode: "129082" },
    { f: "Wingdings 3", code: "183", unicode: "129081" },
    { f: "Wingdings 3", code: "184", unicode: "129083" },
    { f: "Wingdings 3", code: "185", unicode: "129176" },
    { f: "Wingdings 3", code: "186", unicode: "129178" },
    { f: "Wingdings 3", code: "187", unicode: "129177" },
    { f: "Wingdings 3", code: "188", unicode: "129179" },
    { f: "Wingdings 3", code: "189", unicode: "129084" },
    { f: "Wingdings 3", code: "190", unicode: "129086" },
    { f: "Wingdings 3", code: "191", unicode: "129085" },
    { f: "Wingdings 3", code: "192", unicode: "129087" },
    { f: "Wingdings 3", code: "193", unicode: "129088" },
    { f: "Wingdings 3", code: "194", unicode: "129090" },
    { f: "Wingdings 3", code: "195", unicode: "129089" },
    { f: "Wingdings 3", code: "196", unicode: "129091" },
    { f: "Wingdings 3", code: "197", unicode: "129092" },
    { f: "Wingdings 3", code: "198", unicode: "129094" },
    { f: "Wingdings 3", code: "199", unicode: "129093" },
    { f: "Wingdings 3", code: "200", unicode: "129095" },
    { f: "Wingdings 3", code: "201", unicode: "11176" },
    { f: "Wingdings 3", code: "202", unicode: "11177" },
    { f: "Wingdings 3", code: "203", unicode: "11178" },
    { f: "Wingdings 3", code: "204", unicode: "11179" },
    { f: "Wingdings 3", code: "205", unicode: "11180" },
    { f: "Wingdings 3", code: "206", unicode: "11181" },
    { f: "Wingdings 3", code: "207", unicode: "11182" },
    { f: "Wingdings 3", code: "208", unicode: "11183" },
    { f: "Wingdings 3", code: "209", unicode: "129120" },
    { f: "Wingdings 3", code: "210", unicode: "129122" },
    { f: "Wingdings 3", code: "211", unicode: "129121" },
    { f: "Wingdings 3", code: "212", unicode: "129123" },
    { f: "Wingdings 3", code: "213", unicode: "129124" },
    { f: "Wingdings 3", code: "214", unicode: "129125" },
    { f: "Wingdings 3", code: "215", unicode: "129127" },
    { f: "Wingdings 3", code: "216", unicode: "129126" },
    { f: "Wingdings 3", code: "217", unicode: "129136" },
    { f: "Wingdings 3", code: "218", unicode: "129138" },
    { f: "Wingdings 3", code: "219", unicode: "129137" },
    { f: "Wingdings 3", code: "220", unicode: "129139" },
    { f: "Wingdings 3", code: "221", unicode: "129140" },
    { f: "Wingdings 3", code: "222", unicode: "129141" },
    { f: "Wingdings 3", code: "223", unicode: "129143" },
    { f: "Wingdings 3", code: "224", unicode: "129142" },
    { f: "Wingdings 3", code: "225", unicode: "129152" },
    { f: "Wingdings 3", code: "226", unicode: "129154" },
    { f: "Wingdings 3", code: "227", unicode: "129153" },
    { f: "Wingdings 3", code: "228", unicode: "129155" },
    { f: "Wingdings 3", code: "229", unicode: "129156" },
    { f: "Wingdings 3", code: "230", unicode: "129157" },
    { f: "Wingdings 3", code: "231", unicode: "129159" },
    { f: "Wingdings 3", code: "232", unicode: "129158" },
    { f: "Wingdings 3", code: "233", unicode: "129168" },
    { f: "Wingdings 3", code: "234", unicode: "129170" },
    { f: "Wingdings 3", code: "235", unicode: "129169" },
    { f: "Wingdings 3", code: "236", unicode: "129171" },
    { f: "Wingdings 3", code: "237", unicode: "129172" },
    { f: "Wingdings 3", code: "238", unicode: "129174" },
    { f: "Wingdings 3", code: "239", unicode: "129173" },
    { f: "Wingdings 3", code: "240", unicode: "129175" }
];

let order = 1;
function tXml(xml, options = {}) {
    const POS = options.pos || 0;
    const CHAR_LT = '<';
    const CHAR_GT = '>';
    const CHAR_SLASH = '/';
    const CHAR_DASH = '-';
    const CHAR_EXCLAMATION = '!';
    const CHAR_SINGLE_QUOTE = "'";
    const CHAR_DOUBLE_QUOTE = '"';
    const STOP_CHARS = "\n\t>/= ";
    const VOID_ELEMENTS = ['img', 'br', 'input', 'meta', 'link'];
    const CODE_LT = CHAR_LT.charCodeAt(0);
    const CODE_GT = CHAR_GT.charCodeAt(0);
    const CODE_DASH = CHAR_DASH.charCodeAt(0);
    const CODE_SLASH = CHAR_SLASH.charCodeAt(0);
    const CODE_EXCLAMATION = CHAR_EXCLAMATION.charCodeAt(0);
    const CODE_SINGLE_QUOTE = CHAR_SINGLE_QUOTE.charCodeAt(0);
    const CODE_DOUBLE_QUOTE = CHAR_DOUBLE_QUOTE.charCodeAt(0);
    let pos = POS;
    function parseChildren() {
        const children = [];
        while (xml[pos]) {
            const charCode = xml.charCodeAt(pos);
            if (charCode === CODE_LT) {
                const nextCharCode = xml.charCodeAt(pos + 1);
                if (nextCharCode === CODE_SLASH) {
                    pos = xml.indexOf(CHAR_GT, pos);
                    if (pos + 1)
                        pos += 1;
                    return children;
                }
                if (nextCharCode === CODE_EXCLAMATION) {
                    if (xml.charCodeAt(pos + 2) === CODE_DASH) {
                        while (pos !== -1 &&
                            !(xml.charCodeAt(pos) === CODE_GT &&
                                xml.charCodeAt(pos - 1) === CODE_DASH &&
                                xml.charCodeAt(pos - 2) === CODE_DASH)) {
                            pos = xml.indexOf(CHAR_GT, pos + 1);
                        }
                        if (pos === -1)
                            pos = xml.length;
                    }
                    else {
                        pos += 2;
                        while (xml.charCodeAt(pos) !== CODE_GT && xml[pos]) {
                            pos++;
                        }
                    }
                    pos++;
                    continue;
                }
                const node = parseNode();
                children.push(node);
            }
            else {
                const text = parseText();
                if (text.trim().length > 0) {
                    children.push(text);
                }
                pos++;
            }
        }
        return children;
    }
    function parseText() {
        const start = pos;
        pos = xml.indexOf(CHAR_LT, pos) - 1;
        if (pos === -2) {
            pos = xml.length;
        }
        return xml.slice(start, pos + 1);
    }
    function parseTagName() {
        const start = pos;
        while (STOP_CHARS.indexOf(xml[pos]) === -1 && xml[pos]) {
            pos++;
        }
        return xml.slice(start, pos);
    }
    function parseAttributeValue() {
        const quoteChar = xml[pos];
        const start = ++pos;
        pos = xml.indexOf(quoteChar, start);
        return xml.slice(start, pos);
    }
    function findAttributePosition() {
        const pattern = new RegExp(`\\s${options.attrName}\\s*=[\'"]${options.attrValue}[\'"]`);
        const match = pattern.exec(xml);
        return match ? match.index : -1;
    }
    function parseNode() {
        const node = {};
        pos++;
        node.tagName = parseTagName();
        let hasAttributes = false;
        while (xml.charCodeAt(pos) !== CODE_GT && xml[pos]) {
            const charCode = xml.charCodeAt(pos);
            if ((charCode > 64 && charCode < 91) || (charCode > 96 && charCode < 123)) {
                const attrName = parseTagName();
                let attrValue = null;
                let currentCharCode = xml.charCodeAt(pos);
                while (currentCharCode &&
                    currentCharCode !== CODE_SINGLE_QUOTE &&
                    currentCharCode !== CODE_DOUBLE_QUOTE &&
                    !((currentCharCode > 64 && currentCharCode < 91) ||
                        (currentCharCode > 96 && currentCharCode < 123)) &&
                    currentCharCode !== CODE_GT) {
                    pos++;
                    currentCharCode = xml.charCodeAt(pos);
                }
                if (currentCharCode === CODE_SINGLE_QUOTE ||
                    currentCharCode === CODE_DOUBLE_QUOTE) {
                    attrValue = parseAttributeValue();
                    if (pos === -1)
                        return node;
                }
                else {
                    attrValue = null;
                    pos--;
                }
                if (!hasAttributes) {
                    node.attributes = {};
                    hasAttributes = true;
                }
                node.attributes[attrName] = attrValue;
            }
            pos++;
        }
        if (xml.charCodeAt(pos - 1) !== CODE_SLASH) {
            if (node.tagName === 'script') {
                const contentStart = pos + 1;
                pos = xml.indexOf('</script>', pos);
                node.children = [xml.slice(contentStart, pos - 1)];
                pos += 8;
            }
            else if (node.tagName === 'style') {
                const contentStart = pos + 1;
                pos = xml.indexOf('</style>', pos);
                node.children = [xml.slice(contentStart, pos - 1)];
                pos += 7;
            }
            else if (VOID_ELEMENTS.indexOf(node.tagName) === -1) {
                pos++;
                node.children = parseChildren();
            }
            else {
                pos++;
            }
        }
        else {
            pos++;
        }
        return node;
    }
    let result;
    if (options.attrValue !== undefined) {
        options.attrName = options.attrName || 'id';
        result = [];
        let attrPos;
        while ((attrPos = findAttributePosition()) !== -1) {
            pos = xml.lastIndexOf(CHAR_LT, attrPos);
            if (pos !== -1) {
                result.push(parseNode());
            }
            xml = xml.substr(pos);
            pos = 0;
        }
    }
    else {
        result = options.parseNode ? parseNode() : parseChildren();
    }
    if (options.filter) {
        result = tXml.filter(result, options.filter);
    }
    if (options.simplify) {
        result = tXml.simplify(result);
    }
    result.pos = pos;
    return result;
}
tXml.simplify = (nodes) => {
    const result = {};
    if (nodes === undefined) {
        return {};
    }
    if (nodes.length === 1 && typeof nodes[0] === 'string') {
        return nodes[0];
    }
    nodes.forEach((node) => {
        if (typeof node !== 'object') {
            return;
        }
        if (!result[node.tagName]) {
            result[node.tagName] = [];
        }
        const simplified = tXml.simplify(node.children || []);
        result[node.tagName].push(simplified);
        if (typeof simplified === 'object' && simplified !== null) {
            if (node.attributes) {
                simplified.attrs = node.attributes;
            }
            if (simplified.attrs === undefined) {
                simplified.attrs = { order: order };
            }
            else {
                simplified.attrs.order = order;
            }
            order++;
        }
    });
    for (const key in result) {
        if (result[key].length === 1) {
            result[key] = result[key][0];
        }
    }
    return result;
};
tXml.filter = (nodes, filterFn) => {
    const result = [];
    nodes.forEach((node) => {
        if (typeof node === 'object' && filterFn(node)) {
            result.push(node);
        }
        if (node.children) {
            const filtered = tXml.filter(node.children, filterFn);
            result.push(...filtered);
        }
    });
    return result;
};
tXml.stringify = (nodes) => {
    let xmlString = '';
    function processNodes(nodes) {
        if (!nodes)
            return;
        for (const item of nodes) {
            if (typeof item === 'string') {
                xmlString += item.trim();
            }
            else {
                processNode(item);
            }
        }
    }
    function processNode(node) {
        xmlString += `<${node.tagName}`;
        for (const attr in node.attributes) {
            const value = node.attributes[attr];
            if (value === null) {
                xmlString += ` ${attr}`;
            }
            else if (value.indexOf('"') === -1) {
                xmlString += ` ${attr}="${value.trim()}"`;
            }
            else {
                xmlString += ` ${attr}='${value.trim()}'`;
            }
        }
        xmlString += '>';
        processNodes(node.children);
        xmlString += `</${node.tagName}>`;
    }
    processNodes(nodes);
    return xmlString;
};
tXml.toContentString = (node) => {
    if (Array.isArray(node)) {
        let text = '';
        node.forEach((child) => {
            text += ` ${tXml.toContentString(child)}`;
            text = text.trim();
        });
        return text;
    }
    if (typeof node === 'object') {
        return tXml.toContentString(node.children);
    }
    return ` ${node}`;
};
tXml.getElementById = (xml, id, simplify) => {
    const result = tXml(xml, {
        attrValue: id,
        simplify: simplify
    });
    return simplify ? result : result[0];
};
tXml.getElementsByClassName = (xml, className, simplify) => {
    return tXml(xml, {
        attrName: 'class',
        attrValue: `[a-zA-Z0-9-s ]*${className}[a-zA-Z0-9-s ]*`,
        simplify: simplify
    });
};
tXml.parseStream = (source, chunkSize) => {
    if (typeof chunkSize === 'function') {
        chunkSize = 0;
    }
    if (typeof chunkSize === 'string') {
        chunkSize = chunkSize.length + 2;
    }
    if (typeof source === 'string') {
        const fs = require('fs');
        source = fs.createReadStream(source, { start: chunkSize });
        chunkSize = 0;
    }
    let pos = chunkSize;
    let buffer = '';
    source.on('data', (chunk) => {
        buffer += chunk;
        let lastPos = 0;
        while (true) {
            pos = buffer.indexOf('<', pos) + 1;
            const node = tXml(buffer, { pos: pos, parseNode: true });
            pos = node.pos;
            if (pos > buffer.length - 1 || lastPos > pos) {
                if (lastPos) {
                    buffer = buffer.slice(lastPos);
                    pos = 0;
                    lastPos = 0;
                }
                return;
            }
            source.emit('xml', node);
            lastPos = pos;
        }
    });
    return source;
};

const PPTXXmlUtils = (function () {
    function getTextByPathStr(node, pathStr) {
        return getTextByPathList(node, pathStr.trim().split(/\s+/));
    }
    function getTextByPathList(node, path) {
        if (path.constructor !== Array) {
            throw Error("Error of path type! path is not array.");
        }
        if (node === undefined) {
            return undefined;
        }
        let cur = node;
        const l = path.length;
        for (let i = 0; i < l; i++) {
            cur = cur[path[i]];
            if (cur === undefined) {
                return undefined;
            }
        }
        return cur;
    }
    function setTextByPathList(node, path, value) {
        if (path.constructor !== Array) {
            throw Error("Error of path type! path is not array.");
        }
        if (node === undefined) {
            return undefined;
        }
        let obj = node;
        const len = path.length;
        for (let i = 0; i < len; i++) {
            const p = path[i];
            if (obj[p] === undefined) {
                if (i === len - 1) {
                    obj[p] = value;
                }
                else {
                    obj[p] = {};
                }
            }
            obj = obj[p];
        }
        return obj;
    }
    function eachElement(node, doFunction) {
        if (node === undefined) {
            return;
        }
        let result = "";
        const target = node;
        if (target.constructor === Array) {
            let l = target.length;
            for (let i = 0; i < l; i++) {
                result += doFunction(target[i], i);
            }
        }
        else {
            result += doFunction(target, 0);
        }
        return result;
    }
    function angleToDegrees(angle) {
        if (angle == "" || angle == null) {
            return 0;
        }
        return Math.round(angle / 60000);
    }
    function degreesToRadians(degrees) {
        if (degrees == "" || degrees == null || degrees == undefined) {
            return 0;
        }
        return degrees * (Math.PI / 180);
    }
    function escapeHtml(text) {
        let map = {
            '&': '&amp;',
            '<': '&lt;',
            '>': '&gt;',
            '"': '&quot;',
            "'": '&#039;'
        };
        return text.replace(/[&<>"']/g, (m) => map[m]);
    }
    async function readXmlFile(zip, filename, isSlideContent, appVersion) {
        try {
            const zipFile = zip.file(filename);
            if (!zipFile)
                return null;
            let fileContent = zipFile.async ? await zipFile.async("text") : zipFile.asText();
            if (isSlideContent && appVersion <= 12) {
                fileContent = fileContent.replace(/<!\[CDATA\[(.*?)\]\]>/g, '$1');
            }
            let xmlData = tXml(fileContent, { simplify: 1 });
            if (xmlData["?xml"] !== undefined) {
                return xmlData["?xml"];
            }
            else {
                return xmlData;
            }
        }
        catch (e) {
            return null;
        }
    }
    async function getContentTypes(zip, appVersion) {
        let ContentTypesJson = await PPTXXmlUtils.readXmlFile(zip, "[Content_Types].xml", false, appVersion);
        let subObj = ContentTypesJson["Types"]["Override"];
        let slidesLocArray = [];
        let slideLayoutsLocArray = [];
        for (const item of subObj) {
            switch (item["attrs"]["ContentType"]) {
                case "application/vnd.openxmlformats-officedocument.presentationml.slide+xml":
                    slidesLocArray.push(item["attrs"]["PartName"].substr(1));
                    break;
                case "application/vnd.openxmlformats-officedocument.presentationml.slideLayout+xml":
                    slideLayoutsLocArray.push(item["attrs"]["PartName"].substr(1));
                    break;
            }
        }
        return {
            "slides": slidesLocArray,
            "slideLayouts": slideLayoutsLocArray
        };
    }
    async function getSlideSizeAndSetDefaultTextStyle(zip, settings) {
        let app = await PPTXXmlUtils.readXmlFile(zip, "docProps/app.xml");
        app["Properties"]["AppVersion"];
        let rtenObj = {};
        let content = await PPTXXmlUtils.readXmlFile(zip, "ppt/presentation.xml");
        let sldSzAttrs = content["p:presentation"]["p:sldSz"]["attrs"];
        let sldSzWidth = parseInt(sldSzAttrs["cx"]);
        let sldSzHeight = parseInt(sldSzAttrs["cy"]);
        sldSzAttrs["type"];
        const defaultTextStyle = content["p:presentation"]["p:defaultTextStyle"];
        const slideWidth = sldSzWidth * SLIDE_FACTOR$1 + settings.incSlide.width | 0;
        const slideHeight = sldSzHeight * SLIDE_FACTOR$1 + settings.incSlide.height | 0;
        rtenObj = {
            "width": slideWidth,
            "height": slideHeight,
            defaultTextStyle
        };
        return rtenObj;
    }
    function resolveMediaPath(mediaPath, context, basePath) {
        if (mediaPath.startsWith('ppt/')) {
            return mediaPath;
        }
        let resolvedPath = mediaPath;
        let baseDir = '';
        switch (context) {
            case 'slide':
                baseDir = 'ppt/slides/';
                break;
            case 'master':
                baseDir = 'ppt/slideMasters/';
                break;
            case 'layout':
                baseDir = 'ppt/slideLayouts/';
                break;
            default:
                baseDir = basePath || '';
        }
        if (mediaPath.startsWith('../')) {
            resolvedPath = `ppt/${mediaPath.substring(3)}`;
        }
        else if (!mediaPath.includes('/')) {
            resolvedPath = `ppt/media/${mediaPath}`;
        }
        else {
            resolvedPath = baseDir + mediaPath;
        }
        resolvedPath = resolvedPath.replace(/\/+/g, '/');
        if (resolvedPath.startsWith('./')) {
            resolvedPath = resolvedPath.substring(2);
        }
        return resolvedPath;
    }
    function findMediaFile(zip, originalPath, context, basePath) {
        let file = zip.file(originalPath);
        if (file) {
            return file;
        }
        const resolvedPath = resolveMediaPath(originalPath, context, basePath);
        file = zip.file(resolvedPath);
        if (file) {
            return file;
        }
        const alternativePaths = [];
        if (originalPath.includes('media/') || !originalPath.includes('/')) {
            const fileName = originalPath.split('/').pop();
            alternativePaths.push(`ppt/media/${fileName}`, `media/${fileName}`, fileName);
        }
        if (originalPath.includes('embeddings/')) {
            const fileName = originalPath.split('/').pop();
            alternativePaths.push(`ppt/embeddings/${fileName}`, `embeddings/${fileName}`);
        }
        for (const altPath of alternativePaths) {
            file = zip.file(altPath);
            if (file) {
                return file;
            }
        }
        return null;
    }
    function base64ArrayBuffer(arrayBuffer) {
        if (typeof Buffer !== 'undefined' && Buffer.from) {
            return Buffer.from(arrayBuffer).toString('base64');
        }
        const bytes = new Uint8Array(arrayBuffer);
        const byteLength = bytes.byteLength;
        if (typeof btoa === 'function') {
            const CHUNK_SIZE = 0x8000;
            let binary = '';
            for (let i = 0; i < byteLength; i += CHUNK_SIZE) {
                const chunk = bytes.subarray(i, Math.min(i + CHUNK_SIZE, byteLength));
                binary += String.fromCharCode.apply(null, chunk);
            }
            return btoa(binary);
        }
        const encodings = 'ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz0123456789+/';
        const byteRemainder = byteLength % 3;
        const mainLength = byteLength - byteRemainder;
        const parts = [];
        for (let i = 0; i < mainLength; i += 3) {
            const chunk = (bytes[i] << 16) | (bytes[i + 1] << 8) | bytes[i + 2];
            parts.push(encodings[(chunk & 16515072) >> 18] +
                encodings[(chunk & 258048) >> 12] +
                encodings[(chunk & 4032) >> 6] +
                encodings[chunk & 63]);
        }
        if (byteRemainder === 1) {
            const chunk = bytes[mainLength];
            parts.push(`${encodings[(chunk & 252) >> 2]}${encodings[(chunk & 3) << 4]}==`);
        }
        else if (byteRemainder === 2) {
            const chunk = (bytes[mainLength] << 8) | bytes[mainLength + 1];
            parts.push(`${encodings[(chunk & 64512) >> 10]}${encodings[(chunk & 1008) >> 4]}${encodings[(chunk & 15) << 2]}=`);
        }
        return parts.join('');
    }
    function extractFileExtension(filename) {
        return filename.substr((~-filename.lastIndexOf(".") >>> 0) + 2);
    }
    function getMimeType(imgFileExt) {
        let mimeType = "";
        switch (imgFileExt.toLowerCase()) {
            case "jpg":
            case "jpeg":
                mimeType = "image/jpeg";
                break;
            case "png":
                mimeType = "image/png";
                break;
            case "gif":
                mimeType = "image/gif";
                break;
            case "emf":
                mimeType = "image/x-emf";
                break;
            case "wmf":
                mimeType = "image/x-wmf";
                break;
            case "svg":
                mimeType = "image/svg+xml";
                break;
            case "mp4":
                mimeType = "video/mp4";
                break;
            case "webm":
                mimeType = "video/webm";
                break;
            case "ogg":
                mimeType = "video/ogg";
                break;
            case "avi":
                mimeType = "video/avi";
                break;
            case "mpg":
                mimeType = "video/mpg";
                break;
            case "wmv":
                mimeType = "video/wmv";
                break;
            case "mp3":
                mimeType = "audio/mpeg";
                break;
            case "wav":
                mimeType = "audio/wav";
                break;
            case "bmp":
                mimeType = "image/bmp";
                break;
            case "webp":
                mimeType = "image/webp";
                break;
            case "tif":
            case "tiff":
                mimeType = "image/tiff";
                break;
        }
        return mimeType;
    }
    function getPosition(slideSpNode, pNode, slideLayoutSpNode, slideMasterSpNode, sType) {
        let off;
        let x = -1, y = -1;
        if (slideSpNode !== undefined) {
            off = slideSpNode["a:off"]["attrs"];
        }
        if (off === undefined && slideLayoutSpNode !== undefined) {
            off = slideLayoutSpNode["a:off"]["attrs"];
        }
        else if (off === undefined && slideMasterSpNode !== undefined) {
            off = slideMasterSpNode["a:off"]["attrs"];
        }
        let offX = 0, offY = 0;
        if (sType == "group" && pNode !== undefined) {
            const grpXfrmNode = PPTXXmlUtils.getTextByPathList(pNode, ["p:grpSpPr", "a:xfrm"]);
            if (grpXfrmNode !== undefined && grpXfrmNode["a:chOff"] !== undefined && grpXfrmNode["a:chOff"]["attrs"] !== undefined) {
                offX = parseInt(grpXfrmNode["a:chOff"]["attrs"]["x"]) * SLIDE_FACTOR$1;
                offY = parseInt(grpXfrmNode["a:chOff"]["attrs"]["y"]) * SLIDE_FACTOR$1;
                offX = Math.round(offX * 100) / 100;
                offY = Math.round(offY * 100) / 100;
            }
        }
        else if (sType == "group-abs" && pNode !== undefined) {
            const grpXfrmNode = PPTXXmlUtils.getTextByPathList(pNode, ["p:grpSpPr", "a:xfrm"]);
            if (grpXfrmNode !== undefined && grpXfrmNode["a:chOff"] !== undefined && grpXfrmNode["a:chOff"]["attrs"] !== undefined) {
                offX = parseInt(grpXfrmNode["a:chOff"]["attrs"]["x"]) * SLIDE_FACTOR$1;
                offY = parseInt(grpXfrmNode["a:chOff"]["attrs"]["y"]) * SLIDE_FACTOR$1;
                offX = Math.round(offX * 100) / 100;
                offY = Math.round(offY * 100) / 100;
            }
        }
        if (sType == "group-rotate" && pNode !== undefined && pNode["p:grpSpPr"] !== undefined) {
            const xfrmNode = pNode["p:grpSpPr"]["a:xfrm"];
            const chx = parseInt(xfrmNode["a:chOff"]["attrs"]["x"]) * SLIDE_FACTOR$1;
            const chy = parseInt(xfrmNode["a:chOff"]["attrs"]["y"]) * SLIDE_FACTOR$1;
            offX = Math.round(chx * 100) / 100;
            offY = Math.round(chy * 100) / 100;
        }
        if (off === undefined) {
            return "";
        }
        else {
            x = parseInt(off["x"]) * SLIDE_FACTOR$1;
            y = parseInt(off["y"]) * SLIDE_FACTOR$1;
            x = Math.round(x * 100) / 100;
            y = Math.round(y * 100) / 100;
            let finalX = Math.round((x - offX) * 100) / 100;
            let finalY = Math.round((y - offY) * 100) / 100;
            return (isNaN(finalX) || isNaN(finalY)) ? "" : `top:${finalY}px; left:${finalX}px;`;
        }
    }
    function getSize(slideSpNode, slideLayoutSpNode, slideMasterSpNode) {
        let ext = undefined;
        let w = -1, h = -1;
        if (slideSpNode !== undefined) {
            ext = slideSpNode["a:ext"]["attrs"];
        }
        else if (slideLayoutSpNode !== undefined) {
            ext = slideLayoutSpNode["a:ext"]["attrs"];
        }
        else if (slideMasterSpNode !== undefined) {
            ext = slideMasterSpNode["a:ext"]["attrs"];
        }
        if (ext === undefined) {
            return "";
        }
        else {
            w = parseInt(ext["cx"]) * SLIDE_FACTOR$1;
            h = parseInt(ext["cy"]) * SLIDE_FACTOR$1;
            w = Math.round(w * 100) / 100;
            h = Math.round(h * 100) / 100;
            return (isNaN(w) || isNaN(h)) ? "" : `width:${w}px; height:${h}px;`;
        }
    }
    function IsVideoLink(vdoFile) {
        let urlregex = /^(https?|ftp):\/\/([a-zA-Z0-9.-]+(:[a-zA-Z0-9.&%$-]+)*@)*((25[0-5]|2[0-4][0-9]|1[0-9]{2}|[1-9][0-9]?)(\.(25[0-5]|2[0-4][0-9]|1[0-9]{2}|[1-9]?[0-9])){3}|([a-zA-Z0-9-]+\.)*[a-zA-Z0-9-]+\.(com|edu|gov|int|mil|net|org|biz|arpa|info|name|pro|aero|coop|museum|[a-zA-Z]{2}))(:[0-9]+)*(\/($|[a-zA-Z0-9.,?'\\+&%$#=~_-]+))*$/;
        return urlregex.test(vdoFile);
    }
    function convertYouTubeUrl(videoUrl) {
        if (!videoUrl)
            return videoUrl;
        const youtubeIdPatterns = [
            /(?:youtube\.com\/watch\?v=|youtu\.be\/|youtube\.com\/embed\/|youtube\.com\/v\/|youtube\.com\/shorts\/)([a-zA-Z0-9_-]{11})/,
            /youtube\.com\/watch\?.*list=([a-zA-Z0-9_-]+)/,
            /youtube\.com\/playlist\?list=([a-zA-Z0-9_-]+)/
        ];
        const isYouTubeUrl = videoUrl.includes('youtube.com') || videoUrl.includes('youtu.be');
        if (!isYouTubeUrl) {
            return videoUrl;
        }
        for (const pattern of youtubeIdPatterns) {
            const match = videoUrl.match(pattern);
            if (match && match[1]) {
                const videoId = match[1];
                const embedUrl = `https://www.youtube.com/embed/${videoId}?rel=0&modestbranding=1`;
                return embedUrl;
            }
        }
        if (videoUrl.includes('/embed/')) {
            return videoUrl;
        }
        return videoUrl;
    }
    function convertVimeoUrl(videoUrl) {
        if (!videoUrl)
            return videoUrl;
        const isVimeoUrl = videoUrl.includes('vimeo.com');
        if (!isVimeoUrl) {
            return videoUrl;
        }
        const vimeoMatch = videoUrl.match(/vimeo\.com\/(\d+)/);
        if (vimeoMatch && vimeoMatch[1]) {
            const videoId = vimeoMatch[1];
            return `https://player.vimeo.com/video/${videoId}?title=0&byline=0&portrait=0`;
        }
        return videoUrl;
    }
    function convertVideoToEmbed(videoUrl) {
        if (!videoUrl)
            return videoUrl;
        if (videoUrl.includes('youtube.com') || videoUrl.includes('youtu.be')) {
            return convertYouTubeUrl(videoUrl);
        }
        if (videoUrl.includes('vimeo.com')) {
            return convertVimeoUrl(videoUrl);
        }
        return videoUrl;
    }
    return {
        getTextByPathStr: getTextByPathStr,
        getTextByPathList: getTextByPathList,
        setTextByPathList: setTextByPathList,
        eachElement: eachElement,
        angleToDegrees: angleToDegrees,
        degreesToRadians: degreesToRadians,
        escapeHtml: escapeHtml,
        readXmlFile: readXmlFile,
        getContentTypes: getContentTypes,
        getSlideSizeAndSetDefaultTextStyle: getSlideSizeAndSetDefaultTextStyle,
        resolveMediaPath: resolveMediaPath,
        findMediaFile: findMediaFile,
        base64ArrayBuffer: base64ArrayBuffer,
        extractFileExtension,
        getMimeType,
        getPosition,
        getSize,
        IsVideoLink,
        convertYouTubeUrl,
        convertVimeoUrl,
        convertVideoToEmbed,
    };
})();

// This file is autogenerated. It's used to publish ESM to npm.
function _typeof(obj) {
  "@babel/helpers - typeof";

  return _typeof = "function" == typeof Symbol && "symbol" == typeof Symbol.iterator ? function (obj) {
    return typeof obj;
  } : function (obj) {
    return obj && "function" == typeof Symbol && obj.constructor === Symbol && obj !== Symbol.prototype ? "symbol" : typeof obj;
  }, _typeof(obj);
}

// https://github.com/bgrins/TinyColor
// Brian Grinstead, MIT License

var trimLeft = /^\s+/;
var trimRight = /\s+$/;
function tinycolor$2(color, opts) {
  color = color ? color : "";
  opts = opts || {};

  // If input is already a tinycolor, return itself
  if (color instanceof tinycolor$2) {
    return color;
  }
  // If we are called as a function, call using new instead
  if (!(this instanceof tinycolor$2)) {
    return new tinycolor$2(color, opts);
  }
  var rgb = inputToRGB(color);
  this._originalInput = color, this._r = rgb.r, this._g = rgb.g, this._b = rgb.b, this._a = rgb.a, this._roundA = Math.round(100 * this._a) / 100, this._format = opts.format || rgb.format;
  this._gradientType = opts.gradientType;

  // Don't let the range of [0,255] come back in [0,1].
  // Potentially lose a little bit of precision here, but will fix issues where
  // .5 gets interpreted as half of the total, instead of half of 1
  // If it was supposed to be 128, this was already taken care of by `inputToRgb`
  if (this._r < 1) this._r = Math.round(this._r);
  if (this._g < 1) this._g = Math.round(this._g);
  if (this._b < 1) this._b = Math.round(this._b);
  this._ok = rgb.ok;
}
tinycolor$2.prototype = {
  isDark: function isDark() {
    return this.getBrightness() < 128;
  },
  isLight: function isLight() {
    return !this.isDark();
  },
  isValid: function isValid() {
    return this._ok;
  },
  getOriginalInput: function getOriginalInput() {
    return this._originalInput;
  },
  getFormat: function getFormat() {
    return this._format;
  },
  getAlpha: function getAlpha() {
    return this._a;
  },
  getBrightness: function getBrightness() {
    //http://www.w3.org/TR/AERT#color-contrast
    var rgb = this.toRgb();
    return (rgb.r * 299 + rgb.g * 587 + rgb.b * 114) / 1000;
  },
  getLuminance: function getLuminance() {
    //http://www.w3.org/TR/2008/REC-WCAG20-20081211/#relativeluminancedef
    var rgb = this.toRgb();
    var RsRGB, GsRGB, BsRGB, R, G, B;
    RsRGB = rgb.r / 255;
    GsRGB = rgb.g / 255;
    BsRGB = rgb.b / 255;
    if (RsRGB <= 0.03928) R = RsRGB / 12.92;else R = Math.pow((RsRGB + 0.055) / 1.055, 2.4);
    if (GsRGB <= 0.03928) G = GsRGB / 12.92;else G = Math.pow((GsRGB + 0.055) / 1.055, 2.4);
    if (BsRGB <= 0.03928) B = BsRGB / 12.92;else B = Math.pow((BsRGB + 0.055) / 1.055, 2.4);
    return 0.2126 * R + 0.7152 * G + 0.0722 * B;
  },
  setAlpha: function setAlpha(value) {
    this._a = boundAlpha(value);
    this._roundA = Math.round(100 * this._a) / 100;
    return this;
  },
  toHsv: function toHsv() {
    var hsv = rgbToHsv(this._r, this._g, this._b);
    return {
      h: hsv.h * 360,
      s: hsv.s,
      v: hsv.v,
      a: this._a
    };
  },
  toHsvString: function toHsvString() {
    var hsv = rgbToHsv(this._r, this._g, this._b);
    var h = Math.round(hsv.h * 360),
      s = Math.round(hsv.s * 100),
      v = Math.round(hsv.v * 100);
    return this._a == 1 ? "hsv(" + h + ", " + s + "%, " + v + "%)" : "hsva(" + h + ", " + s + "%, " + v + "%, " + this._roundA + ")";
  },
  toHsl: function toHsl() {
    var hsl = rgbToHsl(this._r, this._g, this._b);
    return {
      h: hsl.h * 360,
      s: hsl.s,
      l: hsl.l,
      a: this._a
    };
  },
  toHslString: function toHslString() {
    var hsl = rgbToHsl(this._r, this._g, this._b);
    var h = Math.round(hsl.h * 360),
      s = Math.round(hsl.s * 100),
      l = Math.round(hsl.l * 100);
    return this._a == 1 ? "hsl(" + h + ", " + s + "%, " + l + "%)" : "hsla(" + h + ", " + s + "%, " + l + "%, " + this._roundA + ")";
  },
  toHex: function toHex(allow3Char) {
    return rgbToHex(this._r, this._g, this._b, allow3Char);
  },
  toHexString: function toHexString(allow3Char) {
    return "#" + this.toHex(allow3Char);
  },
  toHex8: function toHex8(allow4Char) {
    return rgbaToHex(this._r, this._g, this._b, this._a, allow4Char);
  },
  toHex8String: function toHex8String(allow4Char) {
    return "#" + this.toHex8(allow4Char);
  },
  toRgb: function toRgb() {
    return {
      r: Math.round(this._r),
      g: Math.round(this._g),
      b: Math.round(this._b),
      a: this._a
    };
  },
  toRgbString: function toRgbString() {
    return this._a == 1 ? "rgb(" + Math.round(this._r) + ", " + Math.round(this._g) + ", " + Math.round(this._b) + ")" : "rgba(" + Math.round(this._r) + ", " + Math.round(this._g) + ", " + Math.round(this._b) + ", " + this._roundA + ")";
  },
  toPercentageRgb: function toPercentageRgb() {
    return {
      r: Math.round(bound01(this._r, 255) * 100) + "%",
      g: Math.round(bound01(this._g, 255) * 100) + "%",
      b: Math.round(bound01(this._b, 255) * 100) + "%",
      a: this._a
    };
  },
  toPercentageRgbString: function toPercentageRgbString() {
    return this._a == 1 ? "rgb(" + Math.round(bound01(this._r, 255) * 100) + "%, " + Math.round(bound01(this._g, 255) * 100) + "%, " + Math.round(bound01(this._b, 255) * 100) + "%)" : "rgba(" + Math.round(bound01(this._r, 255) * 100) + "%, " + Math.round(bound01(this._g, 255) * 100) + "%, " + Math.round(bound01(this._b, 255) * 100) + "%, " + this._roundA + ")";
  },
  toName: function toName() {
    if (this._a === 0) {
      return "transparent";
    }
    if (this._a < 1) {
      return false;
    }
    return hexNames[rgbToHex(this._r, this._g, this._b, true)] || false;
  },
  toFilter: function toFilter(secondColor) {
    var hex8String = "#" + rgbaToArgbHex(this._r, this._g, this._b, this._a);
    var secondHex8String = hex8String;
    var gradientType = this._gradientType ? "GradientType = 1, " : "";
    if (secondColor) {
      var s = tinycolor$2(secondColor);
      secondHex8String = "#" + rgbaToArgbHex(s._r, s._g, s._b, s._a);
    }
    return "progid:DXImageTransform.Microsoft.gradient(" + gradientType + "startColorstr=" + hex8String + ",endColorstr=" + secondHex8String + ")";
  },
  toString: function toString(format) {
    var formatSet = !!format;
    format = format || this._format;
    var formattedString = false;
    var hasAlpha = this._a < 1 && this._a >= 0;
    var needsAlphaFormat = !formatSet && hasAlpha && (format === "hex" || format === "hex6" || format === "hex3" || format === "hex4" || format === "hex8" || format === "name");
    if (needsAlphaFormat) {
      // Special case for "transparent", all other non-alpha formats
      // will return rgba when there is transparency.
      if (format === "name" && this._a === 0) {
        return this.toName();
      }
      return this.toRgbString();
    }
    if (format === "rgb") {
      formattedString = this.toRgbString();
    }
    if (format === "prgb") {
      formattedString = this.toPercentageRgbString();
    }
    if (format === "hex" || format === "hex6") {
      formattedString = this.toHexString();
    }
    if (format === "hex3") {
      formattedString = this.toHexString(true);
    }
    if (format === "hex4") {
      formattedString = this.toHex8String(true);
    }
    if (format === "hex8") {
      formattedString = this.toHex8String();
    }
    if (format === "name") {
      formattedString = this.toName();
    }
    if (format === "hsl") {
      formattedString = this.toHslString();
    }
    if (format === "hsv") {
      formattedString = this.toHsvString();
    }
    return formattedString || this.toHexString();
  },
  clone: function clone() {
    return tinycolor$2(this.toString());
  },
  _applyModification: function _applyModification(fn, args) {
    var color = fn.apply(null, [this].concat([].slice.call(args)));
    this._r = color._r;
    this._g = color._g;
    this._b = color._b;
    this.setAlpha(color._a);
    return this;
  },
  lighten: function lighten() {
    return this._applyModification(_lighten, arguments);
  },
  brighten: function brighten() {
    return this._applyModification(_brighten, arguments);
  },
  darken: function darken() {
    return this._applyModification(_darken, arguments);
  },
  desaturate: function desaturate() {
    return this._applyModification(_desaturate, arguments);
  },
  saturate: function saturate() {
    return this._applyModification(_saturate, arguments);
  },
  greyscale: function greyscale() {
    return this._applyModification(_greyscale, arguments);
  },
  spin: function spin() {
    return this._applyModification(_spin, arguments);
  },
  _applyCombination: function _applyCombination(fn, args) {
    return fn.apply(null, [this].concat([].slice.call(args)));
  },
  analogous: function analogous() {
    return this._applyCombination(_analogous, arguments);
  },
  complement: function complement() {
    return this._applyCombination(_complement, arguments);
  },
  monochromatic: function monochromatic() {
    return this._applyCombination(_monochromatic, arguments);
  },
  splitcomplement: function splitcomplement() {
    return this._applyCombination(_splitcomplement, arguments);
  },
  // Disabled until https://github.com/bgrins/TinyColor/issues/254
  // polyad: function (number) {
  //   return this._applyCombination(polyad, [number]);
  // },
  triad: function triad() {
    return this._applyCombination(polyad, [3]);
  },
  tetrad: function tetrad() {
    return this._applyCombination(polyad, [4]);
  }
};

// If input is an object, force 1 into "1.0" to handle ratios properly
// String input requires "1.0" as input, so 1 will be treated as 1
tinycolor$2.fromRatio = function (color, opts) {
  if (_typeof(color) == "object") {
    var newColor = {};
    for (var i in color) {
      if (color.hasOwnProperty(i)) {
        if (i === "a") {
          newColor[i] = color[i];
        } else {
          newColor[i] = convertToPercentage(color[i]);
        }
      }
    }
    color = newColor;
  }
  return tinycolor$2(color, opts);
};

// Given a string or object, convert that input to RGB
// Possible string inputs:
//
//     "red"
//     "#f00" or "f00"
//     "#ff0000" or "ff0000"
//     "#ff000000" or "ff000000"
//     "rgb 255 0 0" or "rgb (255, 0, 0)"
//     "rgb 1.0 0 0" or "rgb (1, 0, 0)"
//     "rgba (255, 0, 0, 1)" or "rgba 255, 0, 0, 1"
//     "rgba (1.0, 0, 0, 1)" or "rgba 1.0, 0, 0, 1"
//     "hsl(0, 100%, 50%)" or "hsl 0 100% 50%"
//     "hsla(0, 100%, 50%, 1)" or "hsla 0 100% 50%, 1"
//     "hsv(0, 100%, 100%)" or "hsv 0 100% 100%"
//
function inputToRGB(color) {
  var rgb = {
    r: 0,
    g: 0,
    b: 0
  };
  var a = 1;
  var s = null;
  var v = null;
  var l = null;
  var ok = false;
  var format = false;
  if (typeof color == "string") {
    color = stringInputToObject(color);
  }
  if (_typeof(color) == "object") {
    if (isValidCSSUnit(color.r) && isValidCSSUnit(color.g) && isValidCSSUnit(color.b)) {
      rgb = rgbToRgb(color.r, color.g, color.b);
      ok = true;
      format = String(color.r).substr(-1) === "%" ? "prgb" : "rgb";
    } else if (isValidCSSUnit(color.h) && isValidCSSUnit(color.s) && isValidCSSUnit(color.v)) {
      s = convertToPercentage(color.s);
      v = convertToPercentage(color.v);
      rgb = hsvToRgb(color.h, s, v);
      ok = true;
      format = "hsv";
    } else if (isValidCSSUnit(color.h) && isValidCSSUnit(color.s) && isValidCSSUnit(color.l)) {
      s = convertToPercentage(color.s);
      l = convertToPercentage(color.l);
      rgb = hslToRgb$1(color.h, s, l);
      ok = true;
      format = "hsl";
    }
    if (color.hasOwnProperty("a")) {
      a = color.a;
    }
  }
  a = boundAlpha(a);
  return {
    ok: ok,
    format: color.format || format,
    r: Math.min(255, Math.max(rgb.r, 0)),
    g: Math.min(255, Math.max(rgb.g, 0)),
    b: Math.min(255, Math.max(rgb.b, 0)),
    a: a
  };
}

// Conversion Functions
// --------------------

// `rgbToHsl`, `rgbToHsv`, `hslToRgb`, `hsvToRgb` modified from:
// <http://mjijackson.com/2008/02/rgb-to-hsl-and-rgb-to-hsv-color-model-conversion-algorithms-in-javascript>

// `rgbToRgb`
// Handle bounds / percentage checking to conform to CSS color spec
// <http://www.w3.org/TR/css3-color/>
// *Assumes:* r, g, b in [0, 255] or [0, 1]
// *Returns:* { r, g, b } in [0, 255]
function rgbToRgb(r, g, b) {
  return {
    r: bound01(r, 255) * 255,
    g: bound01(g, 255) * 255,
    b: bound01(b, 255) * 255
  };
}

// `rgbToHsl`
// Converts an RGB color value to HSL.
// *Assumes:* r, g, and b are contained in [0, 255] or [0, 1]
// *Returns:* { h, s, l } in [0,1]
function rgbToHsl(r, g, b) {
  r = bound01(r, 255);
  g = bound01(g, 255);
  b = bound01(b, 255);
  var max = Math.max(r, g, b),
    min = Math.min(r, g, b);
  var h,
    s,
    l = (max + min) / 2;
  if (max == min) {
    h = s = 0; // achromatic
  } else {
    var d = max - min;
    s = l > 0.5 ? d / (2 - max - min) : d / (max + min);
    switch (max) {
      case r:
        h = (g - b) / d + (g < b ? 6 : 0);
        break;
      case g:
        h = (b - r) / d + 2;
        break;
      case b:
        h = (r - g) / d + 4;
        break;
    }
    h /= 6;
  }
  return {
    h: h,
    s: s,
    l: l
  };
}

// `hslToRgb`
// Converts an HSL color value to RGB.
// *Assumes:* h is contained in [0, 1] or [0, 360] and s and l are contained [0, 1] or [0, 100]
// *Returns:* { r, g, b } in the set [0, 255]
function hslToRgb$1(h, s, l) {
  var r, g, b;
  h = bound01(h, 360);
  s = bound01(s, 100);
  l = bound01(l, 100);
  function hue2rgb(p, q, t) {
    if (t < 0) t += 1;
    if (t > 1) t -= 1;
    if (t < 1 / 6) return p + (q - p) * 6 * t;
    if (t < 1 / 2) return q;
    if (t < 2 / 3) return p + (q - p) * (2 / 3 - t) * 6;
    return p;
  }
  if (s === 0) {
    r = g = b = l; // achromatic
  } else {
    var q = l < 0.5 ? l * (1 + s) : l + s - l * s;
    var p = 2 * l - q;
    r = hue2rgb(p, q, h + 1 / 3);
    g = hue2rgb(p, q, h);
    b = hue2rgb(p, q, h - 1 / 3);
  }
  return {
    r: r * 255,
    g: g * 255,
    b: b * 255
  };
}

// `rgbToHsv`
// Converts an RGB color value to HSV
// *Assumes:* r, g, and b are contained in the set [0, 255] or [0, 1]
// *Returns:* { h, s, v } in [0,1]
function rgbToHsv(r, g, b) {
  r = bound01(r, 255);
  g = bound01(g, 255);
  b = bound01(b, 255);
  var max = Math.max(r, g, b),
    min = Math.min(r, g, b);
  var h,
    s,
    v = max;
  var d = max - min;
  s = max === 0 ? 0 : d / max;
  if (max == min) {
    h = 0; // achromatic
  } else {
    switch (max) {
      case r:
        h = (g - b) / d + (g < b ? 6 : 0);
        break;
      case g:
        h = (b - r) / d + 2;
        break;
      case b:
        h = (r - g) / d + 4;
        break;
    }
    h /= 6;
  }
  return {
    h: h,
    s: s,
    v: v
  };
}

// `hsvToRgb`
// Converts an HSV color value to RGB.
// *Assumes:* h is contained in [0, 1] or [0, 360] and s and v are contained in [0, 1] or [0, 100]
// *Returns:* { r, g, b } in the set [0, 255]
function hsvToRgb(h, s, v) {
  h = bound01(h, 360) * 6;
  s = bound01(s, 100);
  v = bound01(v, 100);
  var i = Math.floor(h),
    f = h - i,
    p = v * (1 - s),
    q = v * (1 - f * s),
    t = v * (1 - (1 - f) * s),
    mod = i % 6,
    r = [v, q, p, p, t, v][mod],
    g = [t, v, v, q, p, p][mod],
    b = [p, p, t, v, v, q][mod];
  return {
    r: r * 255,
    g: g * 255,
    b: b * 255
  };
}

// `rgbToHex`
// Converts an RGB color to hex
// Assumes r, g, and b are contained in the set [0, 255]
// Returns a 3 or 6 character hex
function rgbToHex(r, g, b, allow3Char) {
  var hex = [pad2(Math.round(r).toString(16)), pad2(Math.round(g).toString(16)), pad2(Math.round(b).toString(16))];

  // Return a 3 character hex if possible
  if (allow3Char && hex[0].charAt(0) == hex[0].charAt(1) && hex[1].charAt(0) == hex[1].charAt(1) && hex[2].charAt(0) == hex[2].charAt(1)) {
    return hex[0].charAt(0) + hex[1].charAt(0) + hex[2].charAt(0);
  }
  return hex.join("");
}

// `rgbaToHex`
// Converts an RGBA color plus alpha transparency to hex
// Assumes r, g, b are contained in the set [0, 255] and
// a in [0, 1]. Returns a 4 or 8 character rgba hex
function rgbaToHex(r, g, b, a, allow4Char) {
  var hex = [pad2(Math.round(r).toString(16)), pad2(Math.round(g).toString(16)), pad2(Math.round(b).toString(16)), pad2(convertDecimalToHex(a))];

  // Return a 4 character hex if possible
  if (allow4Char && hex[0].charAt(0) == hex[0].charAt(1) && hex[1].charAt(0) == hex[1].charAt(1) && hex[2].charAt(0) == hex[2].charAt(1) && hex[3].charAt(0) == hex[3].charAt(1)) {
    return hex[0].charAt(0) + hex[1].charAt(0) + hex[2].charAt(0) + hex[3].charAt(0);
  }
  return hex.join("");
}

// `rgbaToArgbHex`
// Converts an RGBA color to an ARGB Hex8 string
// Rarely used, but required for "toFilter()"
function rgbaToArgbHex(r, g, b, a) {
  var hex = [pad2(convertDecimalToHex(a)), pad2(Math.round(r).toString(16)), pad2(Math.round(g).toString(16)), pad2(Math.round(b).toString(16))];
  return hex.join("");
}

// `equals`
// Can be called with any tinycolor input
tinycolor$2.equals = function (color1, color2) {
  if (!color1 || !color2) return false;
  return tinycolor$2(color1).toRgbString() == tinycolor$2(color2).toRgbString();
};
tinycolor$2.random = function () {
  return tinycolor$2.fromRatio({
    r: Math.random(),
    g: Math.random(),
    b: Math.random()
  });
};

// Modification Functions
// ----------------------
// Thanks to less.js for some of the basics here
// <https://github.com/cloudhead/less.js/blob/master/lib/less/functions.js>

function _desaturate(color, amount) {
  amount = amount === 0 ? 0 : amount || 10;
  var hsl = tinycolor$2(color).toHsl();
  hsl.s -= amount / 100;
  hsl.s = clamp01(hsl.s);
  return tinycolor$2(hsl);
}
function _saturate(color, amount) {
  amount = amount === 0 ? 0 : amount || 10;
  var hsl = tinycolor$2(color).toHsl();
  hsl.s += amount / 100;
  hsl.s = clamp01(hsl.s);
  return tinycolor$2(hsl);
}
function _greyscale(color) {
  return tinycolor$2(color).desaturate(100);
}
function _lighten(color, amount) {
  amount = amount === 0 ? 0 : amount || 10;
  var hsl = tinycolor$2(color).toHsl();
  hsl.l += amount / 100;
  hsl.l = clamp01(hsl.l);
  return tinycolor$2(hsl);
}
function _brighten(color, amount) {
  amount = amount === 0 ? 0 : amount || 10;
  var rgb = tinycolor$2(color).toRgb();
  rgb.r = Math.max(0, Math.min(255, rgb.r - Math.round(255 * -(amount / 100))));
  rgb.g = Math.max(0, Math.min(255, rgb.g - Math.round(255 * -(amount / 100))));
  rgb.b = Math.max(0, Math.min(255, rgb.b - Math.round(255 * -(amount / 100))));
  return tinycolor$2(rgb);
}
function _darken(color, amount) {
  amount = amount === 0 ? 0 : amount || 10;
  var hsl = tinycolor$2(color).toHsl();
  hsl.l -= amount / 100;
  hsl.l = clamp01(hsl.l);
  return tinycolor$2(hsl);
}

// Spin takes a positive or negative amount within [-360, 360] indicating the change of hue.
// Values outside of this range will be wrapped into this range.
function _spin(color, amount) {
  var hsl = tinycolor$2(color).toHsl();
  var hue = (hsl.h + amount) % 360;
  hsl.h = hue < 0 ? 360 + hue : hue;
  return tinycolor$2(hsl);
}

// Combination Functions
// ---------------------
// Thanks to jQuery xColor for some of the ideas behind these
// <https://github.com/infusion/jQuery-xcolor/blob/master/jquery.xcolor.js>

function _complement(color) {
  var hsl = tinycolor$2(color).toHsl();
  hsl.h = (hsl.h + 180) % 360;
  return tinycolor$2(hsl);
}
function polyad(color, number) {
  if (isNaN(number) || number <= 0) {
    throw new Error("Argument to polyad must be a positive number");
  }
  var hsl = tinycolor$2(color).toHsl();
  var result = [tinycolor$2(color)];
  var step = 360 / number;
  for (var i = 1; i < number; i++) {
    result.push(tinycolor$2({
      h: (hsl.h + i * step) % 360,
      s: hsl.s,
      l: hsl.l
    }));
  }
  return result;
}
function _splitcomplement(color) {
  var hsl = tinycolor$2(color).toHsl();
  var h = hsl.h;
  return [tinycolor$2(color), tinycolor$2({
    h: (h + 72) % 360,
    s: hsl.s,
    l: hsl.l
  }), tinycolor$2({
    h: (h + 216) % 360,
    s: hsl.s,
    l: hsl.l
  })];
}
function _analogous(color, results, slices) {
  results = results || 6;
  slices = slices || 30;
  var hsl = tinycolor$2(color).toHsl();
  var part = 360 / slices;
  var ret = [tinycolor$2(color)];
  for (hsl.h = (hsl.h - (part * results >> 1) + 720) % 360; --results;) {
    hsl.h = (hsl.h + part) % 360;
    ret.push(tinycolor$2(hsl));
  }
  return ret;
}
function _monochromatic(color, results) {
  results = results || 6;
  var hsv = tinycolor$2(color).toHsv();
  var h = hsv.h,
    s = hsv.s,
    v = hsv.v;
  var ret = [];
  var modification = 1 / results;
  while (results--) {
    ret.push(tinycolor$2({
      h: h,
      s: s,
      v: v
    }));
    v = (v + modification) % 1;
  }
  return ret;
}

// Utility Functions
// ---------------------

tinycolor$2.mix = function (color1, color2, amount) {
  amount = amount === 0 ? 0 : amount || 50;
  var rgb1 = tinycolor$2(color1).toRgb();
  var rgb2 = tinycolor$2(color2).toRgb();
  var p = amount / 100;
  var rgba = {
    r: (rgb2.r - rgb1.r) * p + rgb1.r,
    g: (rgb2.g - rgb1.g) * p + rgb1.g,
    b: (rgb2.b - rgb1.b) * p + rgb1.b,
    a: (rgb2.a - rgb1.a) * p + rgb1.a
  };
  return tinycolor$2(rgba);
};

// Readability Functions
// ---------------------
// <http://www.w3.org/TR/2008/REC-WCAG20-20081211/#contrast-ratiodef (WCAG Version 2)

// `contrast`
// Analyze the 2 colors and returns the color contrast defined by (WCAG Version 2)
tinycolor$2.readability = function (color1, color2) {
  var c1 = tinycolor$2(color1);
  var c2 = tinycolor$2(color2);
  return (Math.max(c1.getLuminance(), c2.getLuminance()) + 0.05) / (Math.min(c1.getLuminance(), c2.getLuminance()) + 0.05);
};

// `isReadable`
// Ensure that foreground and background color combinations meet WCAG2 guidelines.
// The third argument is an optional Object.
//      the 'level' property states 'AA' or 'AAA' - if missing or invalid, it defaults to 'AA';
//      the 'size' property states 'large' or 'small' - if missing or invalid, it defaults to 'small'.
// If the entire object is absent, isReadable defaults to {level:"AA",size:"small"}.

// *Example*
//    tinycolor.isReadable("#000", "#111") => false
//    tinycolor.isReadable("#000", "#111",{level:"AA",size:"large"}) => false
tinycolor$2.isReadable = function (color1, color2, wcag2) {
  var readability = tinycolor$2.readability(color1, color2);
  var wcag2Parms, out;
  out = false;
  wcag2Parms = validateWCAG2Parms(wcag2);
  switch (wcag2Parms.level + wcag2Parms.size) {
    case "AAsmall":
    case "AAAlarge":
      out = readability >= 4.5;
      break;
    case "AAlarge":
      out = readability >= 3;
      break;
    case "AAAsmall":
      out = readability >= 7;
      break;
  }
  return out;
};

// `mostReadable`
// Given a base color and a list of possible foreground or background
// colors for that base, returns the most readable color.
// Optionally returns Black or White if the most readable color is unreadable.
// *Example*
//    tinycolor.mostReadable(tinycolor.mostReadable("#123", ["#124", "#125"],{includeFallbackColors:false}).toHexString(); // "#112255"
//    tinycolor.mostReadable(tinycolor.mostReadable("#123", ["#124", "#125"],{includeFallbackColors:true}).toHexString();  // "#ffffff"
//    tinycolor.mostReadable("#a8015a", ["#faf3f3"],{includeFallbackColors:true,level:"AAA",size:"large"}).toHexString(); // "#faf3f3"
//    tinycolor.mostReadable("#a8015a", ["#faf3f3"],{includeFallbackColors:true,level:"AAA",size:"small"}).toHexString(); // "#ffffff"
tinycolor$2.mostReadable = function (baseColor, colorList, args) {
  var bestColor = null;
  var bestScore = 0;
  var readability;
  var includeFallbackColors, level, size;
  args = args || {};
  includeFallbackColors = args.includeFallbackColors;
  level = args.level;
  size = args.size;
  for (var i = 0; i < colorList.length; i++) {
    readability = tinycolor$2.readability(baseColor, colorList[i]);
    if (readability > bestScore) {
      bestScore = readability;
      bestColor = tinycolor$2(colorList[i]);
    }
  }
  if (tinycolor$2.isReadable(baseColor, bestColor, {
    level: level,
    size: size
  }) || !includeFallbackColors) {
    return bestColor;
  } else {
    args.includeFallbackColors = false;
    return tinycolor$2.mostReadable(baseColor, ["#fff", "#000"], args);
  }
};

// Big List of Colors
// ------------------
// <https://www.w3.org/TR/css-color-4/#named-colors>
var names = tinycolor$2.names = {
  aliceblue: "f0f8ff",
  antiquewhite: "faebd7",
  aqua: "0ff",
  aquamarine: "7fffd4",
  azure: "f0ffff",
  beige: "f5f5dc",
  bisque: "ffe4c4",
  black: "000",
  blanchedalmond: "ffebcd",
  blue: "00f",
  blueviolet: "8a2be2",
  brown: "a52a2a",
  burlywood: "deb887",
  burntsienna: "ea7e5d",
  cadetblue: "5f9ea0",
  chartreuse: "7fff00",
  chocolate: "d2691e",
  coral: "ff7f50",
  cornflowerblue: "6495ed",
  cornsilk: "fff8dc",
  crimson: "dc143c",
  cyan: "0ff",
  darkblue: "00008b",
  darkcyan: "008b8b",
  darkgoldenrod: "b8860b",
  darkgray: "a9a9a9",
  darkgreen: "006400",
  darkgrey: "a9a9a9",
  darkkhaki: "bdb76b",
  darkmagenta: "8b008b",
  darkolivegreen: "556b2f",
  darkorange: "ff8c00",
  darkorchid: "9932cc",
  darkred: "8b0000",
  darksalmon: "e9967a",
  darkseagreen: "8fbc8f",
  darkslateblue: "483d8b",
  darkslategray: "2f4f4f",
  darkslategrey: "2f4f4f",
  darkturquoise: "00ced1",
  darkviolet: "9400d3",
  deeppink: "ff1493",
  deepskyblue: "00bfff",
  dimgray: "696969",
  dimgrey: "696969",
  dodgerblue: "1e90ff",
  firebrick: "b22222",
  floralwhite: "fffaf0",
  forestgreen: "228b22",
  fuchsia: "f0f",
  gainsboro: "dcdcdc",
  ghostwhite: "f8f8ff",
  gold: "ffd700",
  goldenrod: "daa520",
  gray: "808080",
  green: "008000",
  greenyellow: "adff2f",
  grey: "808080",
  honeydew: "f0fff0",
  hotpink: "ff69b4",
  indianred: "cd5c5c",
  indigo: "4b0082",
  ivory: "fffff0",
  khaki: "f0e68c",
  lavender: "e6e6fa",
  lavenderblush: "fff0f5",
  lawngreen: "7cfc00",
  lemonchiffon: "fffacd",
  lightblue: "add8e6",
  lightcoral: "f08080",
  lightcyan: "e0ffff",
  lightgoldenrodyellow: "fafad2",
  lightgray: "d3d3d3",
  lightgreen: "90ee90",
  lightgrey: "d3d3d3",
  lightpink: "ffb6c1",
  lightsalmon: "ffa07a",
  lightseagreen: "20b2aa",
  lightskyblue: "87cefa",
  lightslategray: "789",
  lightslategrey: "789",
  lightsteelblue: "b0c4de",
  lightyellow: "ffffe0",
  lime: "0f0",
  limegreen: "32cd32",
  linen: "faf0e6",
  magenta: "f0f",
  maroon: "800000",
  mediumaquamarine: "66cdaa",
  mediumblue: "0000cd",
  mediumorchid: "ba55d3",
  mediumpurple: "9370db",
  mediumseagreen: "3cb371",
  mediumslateblue: "7b68ee",
  mediumspringgreen: "00fa9a",
  mediumturquoise: "48d1cc",
  mediumvioletred: "c71585",
  midnightblue: "191970",
  mintcream: "f5fffa",
  mistyrose: "ffe4e1",
  moccasin: "ffe4b5",
  navajowhite: "ffdead",
  navy: "000080",
  oldlace: "fdf5e6",
  olive: "808000",
  olivedrab: "6b8e23",
  orange: "ffa500",
  orangered: "ff4500",
  orchid: "da70d6",
  palegoldenrod: "eee8aa",
  palegreen: "98fb98",
  paleturquoise: "afeeee",
  palevioletred: "db7093",
  papayawhip: "ffefd5",
  peachpuff: "ffdab9",
  peru: "cd853f",
  pink: "ffc0cb",
  plum: "dda0dd",
  powderblue: "b0e0e6",
  purple: "800080",
  rebeccapurple: "663399",
  red: "f00",
  rosybrown: "bc8f8f",
  royalblue: "4169e1",
  saddlebrown: "8b4513",
  salmon: "fa8072",
  sandybrown: "f4a460",
  seagreen: "2e8b57",
  seashell: "fff5ee",
  sienna: "a0522d",
  silver: "c0c0c0",
  skyblue: "87ceeb",
  slateblue: "6a5acd",
  slategray: "708090",
  slategrey: "708090",
  snow: "fffafa",
  springgreen: "00ff7f",
  steelblue: "4682b4",
  tan: "d2b48c",
  teal: "008080",
  thistle: "d8bfd8",
  tomato: "ff6347",
  turquoise: "40e0d0",
  violet: "ee82ee",
  wheat: "f5deb3",
  white: "fff",
  whitesmoke: "f5f5f5",
  yellow: "ff0",
  yellowgreen: "9acd32"
};

// Make it easy to access colors via `hexNames[hex]`
var hexNames = tinycolor$2.hexNames = flip(names);

// Utilities
// ---------

// `{ 'name1': 'val1' }` becomes `{ 'val1': 'name1' }`
function flip(o) {
  var flipped = {};
  for (var i in o) {
    if (o.hasOwnProperty(i)) {
      flipped[o[i]] = i;
    }
  }
  return flipped;
}

// Return a valid alpha value [0,1] with all invalid values being set to 1
function boundAlpha(a) {
  a = parseFloat(a);
  if (isNaN(a) || a < 0 || a > 1) {
    a = 1;
  }
  return a;
}

// Take input from [0, n] and return it as [0, 1]
function bound01(n, max) {
  if (isOnePointZero(n)) n = "100%";
  var processPercent = isPercentage(n);
  n = Math.min(max, Math.max(0, parseFloat(n)));

  // Automatically convert percentage into number
  if (processPercent) {
    n = parseInt(n * max, 10) / 100;
  }

  // Handle floating point rounding errors
  if (Math.abs(n - max) < 0.000001) {
    return 1;
  }

  // Convert into [0, 1] range if it isn't already
  return n % max / parseFloat(max);
}

// Force a number between 0 and 1
function clamp01(val) {
  return Math.min(1, Math.max(0, val));
}

// Parse a base-16 hex value into a base-10 integer
function parseIntFromHex(val) {
  return parseInt(val, 16);
}

// Need to handle 1.0 as 100%, since once it is a number, there is no difference between it and 1
// <http://stackoverflow.com/questions/7422072/javascript-how-to-detect-number-as-a-decimal-including-1-0>
function isOnePointZero(n) {
  return typeof n == "string" && n.indexOf(".") != -1 && parseFloat(n) === 1;
}

// Check to see if string passed in is a percentage
function isPercentage(n) {
  return typeof n === "string" && n.indexOf("%") != -1;
}

// Force a hex value to have 2 characters
function pad2(c) {
  return c.length == 1 ? "0" + c : "" + c;
}

// Replace a decimal with it's percentage value
function convertToPercentage(n) {
  if (n <= 1) {
    n = n * 100 + "%";
  }
  return n;
}

// Converts a decimal to a hex value
function convertDecimalToHex(d) {
  return Math.round(parseFloat(d) * 255).toString(16);
}
// Converts a hex value to a decimal
function convertHexToDecimal(h) {
  return parseIntFromHex(h) / 255;
}
var matchers = function () {
  // <http://www.w3.org/TR/css3-values/#integers>
  var CSS_INTEGER = "[-\\+]?\\d+%?";

  // <http://www.w3.org/TR/css3-values/#number-value>
  var CSS_NUMBER = "[-\\+]?\\d*\\.\\d+%?";

  // Allow positive/negative integer/number.  Don't capture the either/or, just the entire outcome.
  var CSS_UNIT = "(?:" + CSS_NUMBER + ")|(?:" + CSS_INTEGER + ")";

  // Actual matching.
  // Parentheses and commas are optional, but not required.
  // Whitespace can take the place of commas or opening paren
  var PERMISSIVE_MATCH3 = "[\\s|\\(]+(" + CSS_UNIT + ")[,|\\s]+(" + CSS_UNIT + ")[,|\\s]+(" + CSS_UNIT + ")\\s*\\)?";
  var PERMISSIVE_MATCH4 = "[\\s|\\(]+(" + CSS_UNIT + ")[,|\\s]+(" + CSS_UNIT + ")[,|\\s]+(" + CSS_UNIT + ")[,|\\s]+(" + CSS_UNIT + ")\\s*\\)?";
  return {
    CSS_UNIT: new RegExp(CSS_UNIT),
    rgb: new RegExp("rgb" + PERMISSIVE_MATCH3),
    rgba: new RegExp("rgba" + PERMISSIVE_MATCH4),
    hsl: new RegExp("hsl" + PERMISSIVE_MATCH3),
    hsla: new RegExp("hsla" + PERMISSIVE_MATCH4),
    hsv: new RegExp("hsv" + PERMISSIVE_MATCH3),
    hsva: new RegExp("hsva" + PERMISSIVE_MATCH4),
    hex3: /^#?([0-9a-fA-F]{1})([0-9a-fA-F]{1})([0-9a-fA-F]{1})$/,
    hex6: /^#?([0-9a-fA-F]{2})([0-9a-fA-F]{2})([0-9a-fA-F]{2})$/,
    hex4: /^#?([0-9a-fA-F]{1})([0-9a-fA-F]{1})([0-9a-fA-F]{1})([0-9a-fA-F]{1})$/,
    hex8: /^#?([0-9a-fA-F]{2})([0-9a-fA-F]{2})([0-9a-fA-F]{2})([0-9a-fA-F]{2})$/
  };
}();

// `isValidCSSUnit`
// Take in a single string / number and check to see if it looks like a CSS unit
// (see `matchers` above for definition).
function isValidCSSUnit(color) {
  return !!matchers.CSS_UNIT.exec(color);
}

// `stringInputToObject`
// Permissive string parsing.  Take in a number of formats, and output an object
// based on detected format.  Returns `{ r, g, b }` or `{ h, s, l }` or `{ h, s, v}`
function stringInputToObject(color) {
  color = color.replace(trimLeft, "").replace(trimRight, "").toLowerCase();
  var named = false;
  if (names[color]) {
    color = names[color];
    named = true;
  } else if (color == "transparent") {
    return {
      r: 0,
      g: 0,
      b: 0,
      a: 0,
      format: "name"
    };
  }

  // Try to match string input using regular expressions.
  // Keep most of the number bounding out of this function - don't worry about [0,1] or [0,100] or [0,360]
  // Just return an object and let the conversion functions handle that.
  // This way the result will be the same whether the tinycolor is initialized with string or object.
  var match;
  if (match = matchers.rgb.exec(color)) {
    return {
      r: match[1],
      g: match[2],
      b: match[3]
    };
  }
  if (match = matchers.rgba.exec(color)) {
    return {
      r: match[1],
      g: match[2],
      b: match[3],
      a: match[4]
    };
  }
  if (match = matchers.hsl.exec(color)) {
    return {
      h: match[1],
      s: match[2],
      l: match[3]
    };
  }
  if (match = matchers.hsla.exec(color)) {
    return {
      h: match[1],
      s: match[2],
      l: match[3],
      a: match[4]
    };
  }
  if (match = matchers.hsv.exec(color)) {
    return {
      h: match[1],
      s: match[2],
      v: match[3]
    };
  }
  if (match = matchers.hsva.exec(color)) {
    return {
      h: match[1],
      s: match[2],
      v: match[3],
      a: match[4]
    };
  }
  if (match = matchers.hex8.exec(color)) {
    return {
      r: parseIntFromHex(match[1]),
      g: parseIntFromHex(match[2]),
      b: parseIntFromHex(match[3]),
      a: convertHexToDecimal(match[4]),
      format: named ? "name" : "hex8"
    };
  }
  if (match = matchers.hex6.exec(color)) {
    return {
      r: parseIntFromHex(match[1]),
      g: parseIntFromHex(match[2]),
      b: parseIntFromHex(match[3]),
      format: named ? "name" : "hex"
    };
  }
  if (match = matchers.hex4.exec(color)) {
    return {
      r: parseIntFromHex(match[1] + "" + match[1]),
      g: parseIntFromHex(match[2] + "" + match[2]),
      b: parseIntFromHex(match[3] + "" + match[3]),
      a: convertHexToDecimal(match[4] + "" + match[4]),
      format: named ? "name" : "hex8"
    };
  }
  if (match = matchers.hex3.exec(color)) {
    return {
      r: parseIntFromHex(match[1] + "" + match[1]),
      g: parseIntFromHex(match[2] + "" + match[2]),
      b: parseIntFromHex(match[3] + "" + match[3]),
      format: named ? "name" : "hex"
    };
  }
  return false;
}
function validateWCAG2Parms(parms) {
  // return valid WCAG2 parms for isReadable.
  // If input parms are invalid, return {"level":"AA", "size":"small"}
  var level, size;
  parms = parms || {
    level: "AA",
    size: "small"
  };
  level = (parms.level || "AA").toUpperCase();
  size = (parms.size || "small").toLowerCase();
  if (level !== "AA" && level !== "AAA") {
    level = "AA";
  }
  if (size !== "small" && size !== "large") {
    size = "small";
  }
  return {
    level: level,
    size: size
  };
}

const tinycolor$1 = (color, opts) => new tinycolor$2(color, opts);
function getFillType(node) {
    let fillType = "";
    if (node === undefined) {
        return fillType;
    }
    if (node["a:noFill"] !== undefined) {
        fillType = "NO_FILL";
    }
    if (node["a:solidFill"] !== undefined) {
        fillType = "SOLID_FILL";
    }
    if (node["a:gradFill"] !== undefined) {
        fillType = "GRADIENT_FILL";
    }
    if (node["a:pattFill"] !== undefined) {
        fillType = "PATTERN_FILL";
    }
    if (node["a:blipFill"] !== undefined) {
        fillType = "PIC_FILL";
    }
    if (node["a:grpFill"] !== undefined) {
        fillType = "GROUP_FILL";
    }
    return fillType;
}
async function getShapeFill(node, pNode, isSvgMode, warpObj, source) {
    let fillType = getFillType(PPTXXmlUtils.getTextByPathList(node, ["p:spPr"]));
    let fillColor;
    if (fillType === "NO_FILL") {
        return isSvgMode ? "none" : "";
    }
    else if (fillType === "SOLID_FILL") {
        let shpFill = node["p:spPr"]["a:solidFill"];
        fillColor = getSolidFill(shpFill, undefined, undefined, warpObj);
    }
    else if (fillType === "GRADIENT_FILL") {
        let shpFill = node["p:spPr"]["a:gradFill"];
        fillColor = getGradientFill(shpFill, warpObj);
    }
    else if (fillType === "PATTERN_FILL") {
        let shpFill = node["p:spPr"]["a:pattFill"];
        fillColor = getPatternFill(shpFill, warpObj);
    }
    else if (fillType === "PIC_FILL") {
        let shpFill = node["p:spPr"]["a:blipFill"];
        fillColor = await getPicFill(source, shpFill, warpObj);
    }
    if (fillColor === undefined) {
        let clrName = PPTXXmlUtils.getTextByPathList(node, ["p:style", "a:fillRef"]);
        let idx = parseInt(PPTXXmlUtils.getTextByPathList(node, ["p:style", "a:fillRef", "attrs", "idx"]));
        if (idx == 0 || idx == 1000) {
            return isSvgMode ? "none" : "";
        }
        fillColor = getSolidFill(clrName, undefined, undefined, warpObj);
    }
    if (fillColor === undefined) {
        let grpFill = PPTXXmlUtils.getTextByPathList(node, ["p:spPr", "a:grpFill"]);
        if (grpFill !== undefined) {
            let grpShpFill = pNode["p:grpSpPr"];
            let spShpNode = { "p:spPr": grpShpFill };
            return await getShapeFill(spShpNode, node, isSvgMode, warpObj, source);
        }
        else if (fillType === "NO_FILL") {
            return isSvgMode ? "none" : "";
        }
    }
    if (fillColor !== undefined) {
        if (fillType === "GRADIENT_FILL") {
            if (isSvgMode) {
                return fillColor;
            }
            else {
                let { color: colorAry, rot } = fillColor;
                let bgcolor = `background: linear-gradient(${rot}deg,`;
                for (const i of colorAry.keys()) {
                    if (i == colorAry.length - 1) {
                        bgcolor += `#${colorAry[i]});`;
                    }
                    else {
                        bgcolor += `#${colorAry[i]}, `;
                    }
                }
                return bgcolor;
            }
        }
        else if (fillType === "PIC_FILL") {
            if (isSvgMode) {
                if (typeof fillColor === 'object' && fillColor.img) {
                    return fillColor.img;
                }
                else {
                    return fillColor;
                }
            }
            else {
                if (typeof fillColor === 'object' && fillColor.img) {
                    return `background-image:url(${fillColor.img}); background-size: ${fillColor.backgroundSize}; background-position: ${fillColor.backgroundPosition}; background-repeat: ${fillColor.backgroundRepeat};`;
                }
                else {
                    return `background-image:url(${fillColor});`;
                }
            }
        }
        else if (fillType === "PATTERN_FILL") {
            let bgPtrn = "", bgSize = "", bgPos = "";
            bgPtrn = fillColor[0];
            if (fillColor[1] !== null && fillColor[1] !== undefined && fillColor[1] != "") {
                bgSize = ` background-size:${fillColor[1]};`;
            }
            if (fillColor[2] !== null && fillColor[2] !== undefined && fillColor[2] != "") {
                bgPos = ` background-position:${fillColor[2]};`;
            }
            return `background: ${bgPtrn};${bgSize}${bgPos}`;
        }
        else {
            if (isSvgMode) {
                let color = tinycolor$1(fillColor);
                fillColor = color.toRgbString();
                return fillColor;
            }
            else {
                return `background-color: #${fillColor};`;
            }
        }
    }
    else {
        if (isSvgMode) {
            return "none";
        }
        else {
            return "background-color: inherit;";
        }
    }
}
function getFontType(node, type, warpObj, pFontStyle) {
    let typeface = PPTXXmlUtils.getTextByPathList(node, ["a:rPr", "a:latin", "attrs", "typeface"]);
    if (typeface === undefined) {
        let fontIdx = "";
        let fontGrup = "";
        if (pFontStyle !== undefined) {
            fontIdx = PPTXXmlUtils.getTextByPathList(pFontStyle, ["attrs", "idx"]);
        }
        let fontSchemeNode = PPTXXmlUtils.getTextByPathList(warpObj["themeContent"], ["a:theme", "a:themeElements", "a:fontScheme"]);
        if (fontIdx == "") {
            if (type == "title" || type == "subTitle" || type == "ctrTitle") {
                fontIdx = "major";
            }
            else {
                fontIdx = "minor";
            }
        }
        fontGrup = `a:${fontIdx}Font`;
        typeface = PPTXXmlUtils.getTextByPathList(fontSchemeNode, [fontGrup, "a:latin", "attrs", "typeface"]);
    }
    return (typeface === undefined) ? "inherit" : typeface;
}
async function getFontColorPr(node, pNode, lstStyle, pFontStyle, lvl, idx, type, warpObj) {
    let rPrNode = PPTXXmlUtils.getTextByPathList(node, ["a:rPr"]);
    let filTyp, color, textBordr, colorType = "", highlightColor = "";
    if (rPrNode !== undefined) {
        filTyp = getFillType(rPrNode);
        if (filTyp == "SOLID_FILL") {
            let solidFillNode = rPrNode["a:solidFill"];
            color = getSolidFill(solidFillNode, undefined, undefined, warpObj);
            let highlightNode = rPrNode["a:highlight"];
            if (highlightNode !== undefined) {
                highlightColor = getSolidFill(highlightNode, undefined, undefined, warpObj);
            }
            colorType = "solid";
        }
        else if (filTyp == "PATTERN_FILL") {
            let pattFill = rPrNode["a:pattFill"];
            color = getPatternFill(pattFill, warpObj);
            colorType = "pattern";
        }
        else if (filTyp == "PIC_FILL") {
            color = await getBgPicFill(rPrNode, "slideBg", warpObj, undefined);
            colorType = "pic";
        }
        else if (filTyp == "GRADIENT_FILL") {
            let shpFill = rPrNode["a:gradFill"];
            color = getGradientFill(shpFill, warpObj);
            colorType = "gradient";
        }
    }
    if (color === undefined && PPTXXmlUtils.getTextByPathList(lstStyle, [`a:lvl${lvl}pPr`, "a:defRPr"]) !== undefined) {
        let lstStyledefRPr = PPTXXmlUtils.getTextByPathList(lstStyle, [`a:lvl${lvl}pPr`, "a:defRPr"]);
        filTyp = getFillType(lstStyledefRPr);
        if (filTyp == "SOLID_FILL") {
            let solidFillNode = lstStyledefRPr["a:solidFill"];
            color = getSolidFill(solidFillNode, undefined, undefined, warpObj);
            let highlightNode = lstStyledefRPr["a:highlight"];
            if (highlightNode !== undefined) {
                highlightColor = getSolidFill(highlightNode, undefined, undefined, warpObj);
            }
            colorType = "solid";
        }
        else if (filTyp == "PATTERN_FILL") {
            let pattFill = lstStyledefRPr["a:pattFill"];
            color = getPatternFill(pattFill, warpObj);
            colorType = "pattern";
        }
        else if (filTyp == "PIC_FILL") {
            color = await getBgPicFill(lstStyledefRPr, "slideBg", warpObj, undefined);
            colorType = "pic";
        }
        else if (filTyp == "GRADIENT_FILL") {
            let shpFill = lstStyledefRPr["a:gradFill"];
            color = getGradientFill(shpFill, warpObj);
            colorType = "gradient";
        }
    }
    if (color === undefined) {
        let sPstyle = PPTXXmlUtils.getTextByPathList(pNode, ["p:style", "a:fontRef"]);
        if (sPstyle !== undefined) {
            color = getSolidFill(sPstyle, undefined, undefined, warpObj);
            if (color !== undefined) {
                colorType = "solid";
            }
            let highlightNode = sPstyle["a:highlight"];
            if (highlightNode !== undefined) {
                highlightColor = getSolidFill(highlightNode, undefined, undefined, warpObj);
            }
        }
        if (color === undefined) {
            if (pFontStyle !== undefined) {
                color = getSolidFill(pFontStyle, undefined, undefined, warpObj);
                if (color !== undefined) {
                    colorType = "solid";
                }
            }
        }
    }
    if (color === undefined) {
        let layoutMasterNode = getLayoutAndMasterNode(pNode, idx, type, warpObj);
        let { nodeLaout: pPrNodeLaout, nodeMaster: pPrNodeMaster } = layoutMasterNode;
        if (pPrNodeLaout !== undefined) {
            let defRpRLaout = PPTXXmlUtils.getTextByPathList(pPrNodeLaout, ["a:defRPr", "a:solidFill"]);
            if (defRpRLaout !== undefined) {
                color = getSolidFill(defRpRLaout, undefined, undefined, warpObj);
                let highlightNode = PPTXXmlUtils.getTextByPathList(pPrNodeLaout, ["a:defRPr", "a:highlight"]);
                if (highlightNode !== undefined) {
                    highlightColor = getSolidFill(highlightNode, undefined, undefined, warpObj);
                }
                colorType = "solid";
            }
        }
        if (color === undefined) {
            if (pPrNodeMaster !== undefined) {
                let defRprMaster = PPTXXmlUtils.getTextByPathList(pPrNodeMaster, ["a:defRPr", "a:solidFill"]);
                if (defRprMaster !== undefined) {
                    color = getSolidFill(defRprMaster, undefined, undefined, warpObj);
                    let highlightNode = PPTXXmlUtils.getTextByPathList(pPrNodeMaster, ["a:defRPr", "a:highlight"]);
                    if (highlightNode !== undefined) {
                        highlightColor = getSolidFill(highlightNode, undefined, undefined, warpObj);
                    }
                    colorType = "solid";
                }
            }
        }
    }
    let txtEffects = [];
    let txtEffObj = {};
    let txtBrdrNode = PPTXXmlUtils.getTextByPathList(node, ["a:rPr", "a:ln"]);
    textBordr = "";
    if (txtBrdrNode !== undefined && txtBrdrNode["a:noFill"] === undefined) {
        let txBrd = getBorder(node, pNode, false, "text", warpObj);
        let txBrdAry = txBrd.split(" ");
        let brdSize = `${(parseInt(txBrdAry[0].substring(0, txBrdAry[0].indexOf("px"))))}px`;
        let brdClr = txBrdAry[2];
        if (colorType == "solid") {
            textBordr = `-${brdSize} 0 ${brdClr}, 0 ${brdSize} ${brdClr}, ${brdSize} 0 ${brdClr}, 0 -${brdSize} ${brdClr}`;
            txtEffects.push(textBordr);
        }
        else {
            txtEffObj.border = `${brdSize} ${brdClr}`;
        }
    }
    let txtGlowNode = PPTXXmlUtils.getTextByPathList(node, ["a:rPr", "a:effectLst", "a:glow"]);
    let oGlowStr = "";
    if (txtGlowNode !== undefined) {
        let glowClr = getSolidFill(txtGlowNode, undefined, undefined, warpObj);
        let rad = (txtGlowNode["attrs"]["rad"]) ? (txtGlowNode["attrs"]["rad"] * SLIDE_FACTOR$1) : 0;
        oGlowStr = `0 0 ${rad}px #${glowClr}, 0 0 ${rad}px #${glowClr}, 0 0 ${rad}px #${glowClr}, 0 0 ${rad}px #${glowClr}, 0 0 ${rad}px #${glowClr}, 0 0 ${rad}px #${glowClr}, 0 0 ${rad}px #${glowClr}`;
        if (colorType == "solid") {
            txtEffects.push(oGlowStr);
        }
        else {
            txtEffects.push(`drop-shadow(0 0 ${rad / 3}px #${glowClr}) drop-shadow(0 0 ${rad * 2 / 3}px #${glowClr}) drop-shadow(0 0 ${rad}px #${glowClr})`);
        }
    }
    let txtShadow = PPTXXmlUtils.getTextByPathList(node, ["a:rPr", "a:effectLst", "a:outerShdw"]);
    let oShadowStr = "";
    let txtReflection = PPTXXmlUtils.getTextByPathList(node, ["a:rPr", "a:effectLst", "a:reflection"]);
    if (txtReflection !== undefined) {
        parseInt(txtReflection.attrs?.["blurRad"] || "0") * SLIDE_FACTOR$1;
        parseInt(txtReflection.attrs?.["stA"] || "100000");
        parseInt(txtReflection.attrs?.["endA"] || "0");
        parseInt(txtReflection.attrs?.["dist"] || "0") * SLIDE_FACTOR$1;
        parseInt(txtReflection.attrs?.["dir"] || "5400000");
    }
    let txtSoftEdge = PPTXXmlUtils.getTextByPathList(node, ["a:rPr", "a:effectLst", "a:softEdge"]);
    if (txtSoftEdge !== undefined) {
        parseInt(txtSoftEdge.attrs?.["rad"] || "0") * SLIDE_FACTOR$1;
    }
    if (txtShadow === undefined) {
        const effectRefNode = PPTXXmlUtils.getTextByPathList(pNode, ["p:style", "a:effectRef"]);
        if (effectRefNode !== undefined) {
            const effectIdx = PPTXXmlUtils.getTextByPathList(effectRefNode, ["attrs", "idx"]);
            if (effectIdx !== undefined && warpObj["themeContent"] !== undefined) {
                let effectStyleLst = PPTXXmlUtils.getTextByPathList(warpObj["themeContent"], ["a:theme", "a:themeElements", "a:fmtScheme", "a:effectStyleLst", "a:effectStyle"]);
                if (effectStyleLst !== undefined) {
                    if (!Array.isArray(effectStyleLst)) {
                        effectStyleLst = [effectStyleLst];
                    }
                    var idx = Number(effectIdx);
                    if (idx >= 0 && effectStyleLst[idx] !== undefined) {
                        txtShadow = PPTXXmlUtils.getTextByPathList(effectStyleLst[idx], ["a:effectLst", "a:outerShdw"]);
                    }
                }
            }
        }
    }
    if (txtShadow !== undefined) {
        let shadowClr = getSolidFill(txtShadow, undefined, undefined, warpObj);
        let outerShdwAttrs = txtShadow["attrs"];
        outerShdwAttrs["algn"];
        let dir = (outerShdwAttrs["dir"]) ? (parseInt(outerShdwAttrs["dir"]) / 60000) : 0;
        let dist = parseInt(outerShdwAttrs["dist"]) * SLIDE_FACTOR$1;
        outerShdwAttrs["rotWithShape"];
        let blurRad = (outerShdwAttrs["blurRad"]) ? (`${parseInt(outerShdwAttrs["blurRad"]) * SLIDE_FACTOR$1}px`) : "";
        (outerShdwAttrs["sx"]) ? (parseInt(outerShdwAttrs["sx"]) / 100000) : 1;
        (outerShdwAttrs["sy"]) ? (parseInt(outerShdwAttrs["sy"]) / 100000) : 1;
        let vx = dist * Math.sin(dir * Math.PI / 180);
        let hx = dist * Math.cos(dir * Math.PI / 180);
        if (!isNaN(vx) && !isNaN(hx)) {
            oShadowStr = `${hx}px ${vx}px ${blurRad} #${shadowClr}`;
            if (colorType == "solid") {
                txtEffects.push(oShadowStr);
            }
            else {
                txtEffects.push(`drop-shadow(${hx}px ${vx}px ${blurRad} #${shadowClr})`);
            }
        }
    }
    let text_effcts = "", txt_effects;
    if (colorType == "solid") {
        if (txtEffects.length > 0) {
            text_effcts = txtEffects.join(",");
        }
        txt_effects = `${text_effcts};`;
    }
    else {
        if (txtEffects.length > 0) {
            text_effcts = txtEffects.join(" ");
        }
        txtEffObj.effcts = text_effcts;
        txt_effects = txtEffObj;
    }
    return [color, txt_effects, colorType, highlightColor];
}
function getFontSize(node, textBodyNode, pFontStyle, lvl, type, warpObj) {
    let lstStyle = (textBodyNode !== undefined) ? textBodyNode["a:lstStyle"] : undefined;
    let lvlpPr = `a:lvl${lvl}pPr`;
    let fontSize = undefined;
    let sz, kern;
    if (node["a:rPr"] !== undefined && node["a:rPr"]["attrs"] && node["a:rPr"]["attrs"]["sz"] !== undefined) {
        fontSize = parseInt(node["a:rPr"]["attrs"]["sz"]) / 100;
    }
    if (isNaN(fontSize) || fontSize === undefined && node["a:fld"] !== undefined) {
        sz = PPTXXmlUtils.getTextByPathList(node["a:fld"], ["a:rPr", "attrs", "sz"]);
        fontSize = parseInt(sz) / 100;
    }
    if ((isNaN(fontSize) || fontSize === undefined) && node["a:t"] === undefined) {
        sz = PPTXXmlUtils.getTextByPathList(node["a:endParaRPr"], ["attrs", "sz"]);
        fontSize = parseInt(sz) / 100;
    }
    if ((isNaN(fontSize) || fontSize === undefined) && lstStyle !== undefined) {
        sz = PPTXXmlUtils.getTextByPathList(lstStyle, [lvlpPr, "a:defRPr", "attrs", "sz"]);
        fontSize = parseInt(sz) / 100;
    }
    if (textBodyNode !== undefined) {
        PPTXXmlUtils.getTextByPathList(textBodyNode, ["a:bodyPr", "a:spAutoFit"]);
    }
    if (isNaN(fontSize) || fontSize === undefined) {
        sz = PPTXXmlUtils.getTextByPathList(warpObj["slideLayoutTables"], ["typeTable", type, "p:txBody", "a:lstStyle", lvlpPr, "a:defRPr", "attrs", "sz"]);
        fontSize = parseInt(sz) / 100;
        kern = PPTXXmlUtils.getTextByPathList(warpObj["slideLayoutTables"], ["typeTable", type, "p:txBody", "a:lstStyle", lvlpPr, "a:defRPr", "attrs", "kern"]);
    }
    if (isNaN(fontSize) || fontSize === undefined) {
        sz = PPTXXmlUtils.getTextByPathList(warpObj["slideMasterTables"], ["typeTable", type, "p:txBody", "a:lstStyle", lvlpPr, "a:defRPr", "attrs", "sz"]);
        kern = PPTXXmlUtils.getTextByPathList(warpObj["slideMasterTables"], ["typeTable", type, "p:txBody", "a:lstStyle", lvlpPr, "a:defRPr", "attrs", "kern"]);
        if (sz === undefined) {
            if (type == "title" || type == "subTitle" || type == "ctrTitle") {
                sz = PPTXXmlUtils.getTextByPathList(warpObj["slideMasterTextStyles"], ["p:titleStyle", lvlpPr, "a:defRPr", "attrs", "sz"]);
                kern = PPTXXmlUtils.getTextByPathList(warpObj["slideMasterTextStyles"], ["p:titleStyle", lvlpPr, "a:defRPr", "attrs", "kern"]);
            }
            else if (type == "body" || type == "obj" || type == "dt" || type == "sldNum" || type === "textBox") {
                sz = PPTXXmlUtils.getTextByPathList(warpObj["slideMasterTextStyles"], ["p:bodyStyle", lvlpPr, "a:defRPr", "attrs", "sz"]);
                kern = PPTXXmlUtils.getTextByPathList(warpObj["slideMasterTextStyles"], ["p:bodyStyle", lvlpPr, "a:defRPr", "attrs", "kern"]);
            }
            else if (type == "shape") {
                sz = PPTXXmlUtils.getTextByPathList(warpObj["slideMasterTextStyles"], ["p:otherStyle", lvlpPr, "a:defRPr", "attrs", "sz"]);
                kern = PPTXXmlUtils.getTextByPathList(warpObj["slideMasterTextStyles"], ["p:otherStyle", lvlpPr, "a:defRPr", "attrs", "kern"]);
            }
            if (sz === undefined) {
                sz = PPTXXmlUtils.getTextByPathList(warpObj["defaultTextStyle"], [lvlpPr, "a:defRPr", "attrs", "sz"]);
                kern = (kern === undefined) ? PPTXXmlUtils.getTextByPathList(warpObj["defaultTextStyle"], [lvlpPr, "a:defRPr", "attrs", "kern"]) : undefined;
            }
        }
        fontSize = parseInt(sz) / 100;
    }
    let baseline = PPTXXmlUtils.getTextByPathList(node, ["a:rPr", "attrs", "baseline"]);
    if (baseline !== undefined && !isNaN(fontSize)) {
        let baselineVl = parseInt(baseline) / 100000;
        fontSize -= baselineVl;
    }
    if (isNaN(fontSize) || fontSize === undefined) {
        let pPrNode = node.parentNode && node.parentNode["a:pPr"];
        if (pPrNode) {
            let defRPrNode = PPTXXmlUtils.getTextByPathList(pPrNode, ["a:defRPr"]);
            if (defRPrNode && defRPrNode["attrs"] && defRPrNode["attrs"]["sz"]) {
                fontSize = parseInt(defRPrNode["attrs"]["sz"]) / 100;
            }
        }
    }
    if (isNaN(fontSize) || fontSize === undefined) {
        fontSize = 18;
    }
    if (!isNaN(fontSize)) {
        let normAutofit = PPTXXmlUtils.getTextByPathList(textBodyNode, ["a:bodyPr", "a:normAutofit", "attrs", "fontScale"]);
        if (normAutofit !== undefined && normAutofit != 0) {
            fontSize = Math.round(fontSize * (normAutofit / 100000));
        }
    }
    return isNaN(fontSize) ? ((type == "br") ? "initial" : "inherit") : (`${fontSize * FONT_SIZE_FACTOR}px`);
}
function getFontBold(node, type, slideMasterTextStyles) {
    if (node["a:rPr"] !== undefined && node["a:rPr"]["attrs"] !== undefined) {
        const boldAttr = node["a:rPr"]["attrs"]["b"];
        return (boldAttr === "1" || boldAttr === "true" || boldAttr === "on") ? "bold" : "inherit";
    }
    return "inherit";
}
function getFontItalic(node, type, slideMasterTextStyles) {
    return (node["a:rPr"] !== undefined && node["a:rPr"]["attrs"]["i"] === "1") ? "italic" : "inherit";
}
function getFontDecoration(node, type, slideMasterTextStyles) {
    if (node["a:rPr"] !== undefined) {
        let underLine = node["a:rPr"]["attrs"]["u"] !== undefined ? node["a:rPr"]["attrs"]["u"] : "none";
        let strikethrough = node["a:rPr"]["attrs"]["strike"] !== undefined ? node["a:rPr"]["attrs"]["strike"] : 'noStrike';
        if (underLine != "none" && strikethrough == "noStrike") {
            return "underline";
        }
        else if (underLine == "none" && strikethrough != "noStrike") {
            return "line-through";
        }
        else if (underLine != "none" && strikethrough != "noStrike") {
            return "underline line-through";
        }
        else {
            return "inherit";
        }
    }
    else {
        return "inherit";
    }
}
function getTextHorizontalAlign(node, pNode, type, warpObj) {
    let getAlgn = PPTXXmlUtils.getTextByPathList(node, ["a:pPr", "attrs", "algn"]);
    if (getAlgn === undefined) {
        getAlgn = PPTXXmlUtils.getTextByPathList(pNode, ["a:pPr", "attrs", "algn"]);
    }
    if (getAlgn === undefined) {
        if (type == "title" || type == "ctrTitle" || type == "subTitle") {
            let lvlIdx = 1;
            let lvlNode = PPTXXmlUtils.getTextByPathList(pNode, ["a:pPr", "attrs", "lvl"]);
            if (lvlNode !== undefined) {
                lvlIdx = parseInt(lvlNode) + 1;
            }
            let lvlStr = `a:lvl${lvlIdx}pPr`;
            getAlgn = PPTXXmlUtils.getTextByPathList(warpObj, ["slideLayoutTables", "typeTable", type, "p:txBody", "a:lstStyle", lvlStr, "attrs", "algn"]);
            if (getAlgn === undefined) {
                getAlgn = PPTXXmlUtils.getTextByPathList(warpObj, ["slideMasterTables", "typeTable", type, "p:txBody", "a:lstStyle", lvlStr, "attrs", "algn"]);
                if (getAlgn === undefined) {
                    getAlgn = PPTXXmlUtils.getTextByPathList(warpObj, ["slideMasterTextStyles", "p:titleStyle", lvlStr, "attrs", "algn"]);
                    if (getAlgn === undefined && type === "subTitle") {
                        getAlgn = PPTXXmlUtils.getTextByPathList(warpObj, ["slideMasterTextStyles", "p:bodyStyle", lvlStr, "attrs", "algn"]);
                    }
                }
            }
        }
        else if (type == "body") {
            getAlgn = PPTXXmlUtils.getTextByPathList(warpObj, ["slideMasterTextStyles", "p:bodyStyle", "a:lvl1pPr", "attrs", "algn"]);
        }
        else {
            getAlgn = PPTXXmlUtils.getTextByPathList(warpObj, ["slideMasterTables", "typeTable", type, "p:txBody", "a:lstStyle", "a:lvl1pPr", "attrs", "algn"]);
        }
    }
    let align = "inherit";
    if (getAlgn !== undefined) {
        switch (getAlgn) {
            case "l":
                align = "left";
                break;
            case "r":
                align = "right";
                break;
            case "ctr":
                align = "center";
                break;
            case "just":
                align = "justify";
                break;
            case "dist":
                align = "justify";
                break;
            default:
                align = "inherit";
        }
    }
    return align;
}
function getTextVerticalAlign(node, type, slideMasterTextStyles) {
    let baseline = PPTXXmlUtils.getTextByPathList(node, ["a:rPr", "attrs", "baseline"]);
    return baseline === undefined ? "baseline" : `${(parseInt(baseline) / 1000)}%`;
}
function getTableBorders(node, warpObj) {
    let borderStyle = "";
    if (node["a:bottom"] !== undefined) {
        let obj = {
            "p:spPr": {
                "a:ln": node["a:bottom"]["a:ln"]
            }
        };
        let borders = getBorder(obj, undefined, false, "shape", warpObj);
        borderStyle += borders.replace("border", "border-bottom");
    }
    if (node["a:top"] !== undefined) {
        let obj = {
            "p:spPr": {
                "a:ln": node["a:top"]["a:ln"]
            }
        };
        let borders = getBorder(obj, undefined, false, "shape", warpObj);
        borderStyle += borders.replace("border", "border-top");
    }
    if (node["a:right"] !== undefined) {
        let obj = {
            "p:spPr": {
                "a:ln": node["a:right"]["a:ln"]
            }
        };
        let borders = getBorder(obj, undefined, false, "shape", warpObj);
        borderStyle += borders.replace("border", "border-right");
    }
    if (node["a:left"] !== undefined) {
        let obj = {
            "p:spPr": {
                "a:ln": node["a:left"]["a:ln"]
            }
        };
        let borders = getBorder(obj, undefined, false, "shape", warpObj);
        borderStyle += borders.replace("border", "border-left");
    }
    return borderStyle;
}
function getBorder(node, pNode, isSvgMode, bType, warpObj) {
    let cssText, lineNode;
    if (bType == "shape") {
        cssText = "border: ";
        lineNode = node["p:spPr"]["a:ln"];
    }
    else if (bType == "text") {
        cssText = "";
        lineNode = node["a:rPr"]["a:ln"];
    }
    let is_noFill = PPTXXmlUtils.getTextByPathList(lineNode, ["a:noFill"]);
    if (is_noFill !== undefined) {
        return "hidden";
    }
    let lnRefNode;
    let phClr = undefined;
    if (lineNode == undefined) {
        lnRefNode = PPTXXmlUtils.getTextByPathList(node, ["p:style", "a:lnRef"]);
        if (lnRefNode !== undefined) {
            let lnIdx = PPTXXmlUtils.getTextByPathList(lnRefNode, ["attrs", "idx"]);
            if (lnRefNode !== undefined) {
                phClr = getSolidFill(lnRefNode, undefined, undefined, warpObj);
            }
            const lnStyleLst = warpObj["themeContent"]["a:theme"]["a:themeElements"]["a:fmtScheme"]["a:lnStyleLst"]["a:ln"];
            if (Array.isArray(lnStyleLst)) {
                lineNode = lnStyleLst[Number(lnIdx)];
            }
            else {
                lineNode = lnStyleLst;
            }
        }
    }
    if (lineNode == undefined) {
        cssText = "";
        lineNode = node;
    }
    let borderColor;
    let borderWidth = 0;
    let borderType = "solid";
    let strokeDasharray = "0";
    if (lineNode !== undefined) {
        let w = PPTXXmlUtils.getTextByPathList(lineNode, ["attrs", "w"]);
        borderWidth = (w !== undefined) ? parseInt(w) / 12700 : (4 / 3);
        if (isNaN(borderWidth) || borderWidth < 1) {
            cssText += `${(4 / 3)}px `;
        }
        else {
            cssText += `${borderWidth}px `;
        }
        borderType = PPTXXmlUtils.getTextByPathList(lineNode, ["a:prstDash", "attrs", "val"]);
        if (borderType === undefined) {
            borderType = PPTXXmlUtils.getTextByPathList(lineNode, ["attrs", "cmpd"]);
        }
        strokeDasharray = "0";
        switch (borderType) {
            case "solid":
                cssText += "solid";
                strokeDasharray = "0";
                break;
            case "dash":
                cssText += "dashed";
                strokeDasharray = "5";
                break;
            case "dashDot":
                cssText += "dashed";
                strokeDasharray = "5, 5, 1, 5";
                break;
            case "dot":
                cssText += "dotted";
                strokeDasharray = "1, 5";
                break;
            case "lgDash":
                cssText += "dashed";
                strokeDasharray = "10, 5";
                break;
            case "dbl":
                cssText += "double";
                strokeDasharray = "0";
                break;
            case "lgDashDotDot":
                cssText += "dashed";
                strokeDasharray = "10, 5, 1, 5, 1, 5";
                break;
            case "sysDash":
                cssText += "dashed";
                strokeDasharray = "5, 2";
                break;
            case "sysDashDot":
                cssText += "dashed";
                strokeDasharray = "5, 2, 1, 5";
                break;
            case "sysDashDotDot":
                cssText += "dashed";
                strokeDasharray = "5, 2, 1, 5, 1, 5";
                break;
            case "sysDot":
                cssText += "dotted";
                strokeDasharray = "2, 5";
                break;
            case undefined:
            default:
                cssText += "solid";
                strokeDasharray = "0";
        }
        let fillTyp = getFillType(lineNode);
        if (fillTyp === "NO_FILL") {
            borderColor = isSvgMode ? "none" : "";
        }
        else if (fillTyp === "SOLID_FILL") {
            if (!lnRefNode) {
                lnRefNode = PPTXXmlUtils.getTextByPathList(node, ["p:style", "a:lnRef"]);
            }
            if (phClr === undefined && lnRefNode !== undefined) {
                phClr = getSolidFill(lnRefNode, undefined, undefined, warpObj);
            }
            borderColor = getSolidFill(lineNode["a:solidFill"], undefined, phClr, warpObj);
        }
        else if (fillTyp === "GRADIENT_FILL") {
            borderColor = getGradientFill(lineNode["a:gradFill"], warpObj);
        }
        else if (fillTyp === "PATTERN_FILL") {
            borderColor = getPatternFill(lineNode["a:pattFill"], warpObj);
        }
    }
    if (borderColor === undefined) {
        let lnRefNode = PPTXXmlUtils.getTextByPathList(node, ["p:style", "a:lnRef"]);
        if (lnRefNode !== undefined) {
            borderColor = getSolidFill(lnRefNode, undefined, undefined, warpObj);
        }
    }
    if (borderColor === undefined) {
        if (isSvgMode) {
            borderColor = "none";
        }
        else {
            borderColor = "hidden";
        }
    }
    else {
        if (borderColor && typeof borderColor === 'string') {
            if (!borderColor.startsWith('#') && !borderColor.startsWith('rgb') && !borderColor.startsWith('hsl') && borderColor !== 'none' && borderColor !== 'hidden') {
                borderColor = `#${borderColor}`;
            }
        }
    }
    cssText += ` ${borderColor} `;
    if (isSvgMode) {
        let result = { "color": borderColor, "width": borderWidth, "type": borderType, "strokeDasharray": strokeDasharray };
        return result;
    }
    else {
        return `${cssText};`;
    }
}
async function getSlideBackgroundFill(warpObj, index) {
    let slideContent = warpObj["slideContent"];
    let slideLayoutContent = warpObj["slideLayoutContent"];
    let slideMasterContent = warpObj["slideMasterContent"];
    let bgPr = PPTXXmlUtils.getTextByPathList(slideContent, ["p:sld", "p:cSld", "p:bg", "p:bgPr"]);
    let bgRef = PPTXXmlUtils.getTextByPathList(slideContent, ["p:sld", "p:cSld", "p:bg", "p:bgRef"]);
    let bgcolor;
    if (bgPr !== undefined) {
        let bgFillTyp = getFillType(bgPr);
        if (bgFillTyp === "SOLID_FILL") {
            let sldFill = bgPr["a:solidFill"];
            let clrMapOvr;
            let sldClrMapOvr = PPTXXmlUtils.getTextByPathList(slideContent, ["p:sld", "p:clrMapOvr", "a:overrideClrMapping", "attrs"]);
            if (sldClrMapOvr !== undefined) {
                clrMapOvr = sldClrMapOvr;
            }
            else {
                let sldClrMapOvr = PPTXXmlUtils.getTextByPathList(slideLayoutContent, ["p:sldLayout", "p:clrMapOvr", "a:overrideClrMapping", "attrs"]);
                if (sldClrMapOvr !== undefined) {
                    clrMapOvr = sldClrMapOvr;
                }
                else {
                    clrMapOvr = PPTXXmlUtils.getTextByPathList(slideMasterContent, ["p:sldMaster", "p:clrMap", "attrs"]);
                }
            }
            let sldBgClr = getSolidFill(sldFill, clrMapOvr, undefined, warpObj);
            bgcolor = `background: #${sldBgClr};`;
        }
        else if (bgFillTyp === "GRADIENT_FILL") {
            bgcolor = getBgGradientFill(bgPr, undefined, slideMasterContent, warpObj);
        }
        else if (bgFillTyp === "PIC_FILL") {
            bgcolor = await getBgPicFill(bgPr, "slideBg", warpObj, undefined);
        }
    }
    else if (bgRef !== undefined) {
        let clrMapOvr;
        let sldClrMapOvr = PPTXXmlUtils.getTextByPathList(slideContent, ["p:sld", "p:clrMapOvr", "a:overrideClrMapping", "attrs"]);
        if (sldClrMapOvr !== undefined) {
            clrMapOvr = sldClrMapOvr;
        }
        else {
            let sldClrMapOvr = PPTXXmlUtils.getTextByPathList(slideLayoutContent, ["p:sldLayout", "p:clrMapOvr", "a:overrideClrMapping", "attrs"]);
            if (sldClrMapOvr !== undefined) {
                clrMapOvr = sldClrMapOvr;
            }
            else {
                clrMapOvr = PPTXXmlUtils.getTextByPathList(slideMasterContent, ["p:sldMaster", "p:clrMap", "attrs"]);
            }
        }
        let phClr = getSolidFill(bgRef, clrMapOvr, undefined, warpObj);
        let idx = Number(bgRef["attrs"]["idx"]);
        if (idx == 0 || idx == 1000) ;
        else if (idx > 0 && idx < 1000) ;
        else if (idx > 1000) {
            let trueIdx = idx - 1000;
            let bgFillLst = warpObj["themeContent"]["a:theme"]["a:themeElements"]["a:fmtScheme"]["a:bgFillStyleLst"];
            let sortblAry = [];
            Object.keys(bgFillLst).forEach(key => {
                let bgFillLstTyp = bgFillLst[key];
                if (key != "attrs") {
                    if (bgFillLstTyp.constructor === Array) {
                        for (const item of bgFillLstTyp) {
                            let obj = {};
                            obj[key] = item;
                            obj["idex"] = item["attrs"]["order"];
                            obj["attrs"] = {
                                "order": item["attrs"]["order"]
                            };
                            sortblAry.push(obj);
                        }
                    }
                    else {
                        let obj = {};
                        obj[key] = bgFillLstTyp;
                        obj["idex"] = bgFillLstTyp["attrs"]["order"];
                        obj["attrs"] = {
                            "order": bgFillLstTyp["attrs"]["order"]
                        };
                        sortblAry.push(obj);
                    }
                }
            });
            let sortByOrder = sortblAry.slice(0);
            sortByOrder.sort((a, b) => {
                return a.idex - b.idex;
            });
            let bgFillLstIdx = sortByOrder[trueIdx - 1];
            let bgFillTyp = getFillType(bgFillLstIdx);
            if (bgFillTyp === "SOLID_FILL") {
                let sldFill = bgFillLstIdx["a:solidFill"];
                let sldBgClr = getSolidFill(sldFill, clrMapOvr, undefined, warpObj);
                bgcolor = `background: #${sldBgClr};`;
            }
            else if (bgFillTyp === "GRADIENT_FILL") {
                bgcolor = getBgGradientFill(bgFillLstIdx, phClr, slideMasterContent, warpObj);
            }
            else ;
        }
    }
    else {
        bgPr = PPTXXmlUtils.getTextByPathList(slideLayoutContent, ["p:sldLayout", "p:cSld", "p:bg", "p:bgPr"]);
        bgRef = PPTXXmlUtils.getTextByPathList(slideLayoutContent, ["p:sldLayout", "p:cSld", "p:bg", "p:bgRef"]);
        let clrMapOvr;
        let sldClrMapOvr = PPTXXmlUtils.getTextByPathList(slideLayoutContent, ["p:sldLayout", "p:clrMapOvr", "a:overrideClrMapping", "attrs"]);
        if (sldClrMapOvr !== undefined) {
            clrMapOvr = sldClrMapOvr;
        }
        else {
            clrMapOvr = PPTXXmlUtils.getTextByPathList(slideMasterContent, ["p:sldMaster", "p:clrMap", "attrs"]);
        }
        if (bgPr !== undefined) {
            let bgFillTyp = getFillType(bgPr);
            if (bgFillTyp === "SOLID_FILL") {
                let sldFill = bgPr["a:solidFill"];
                let sldBgClr = getSolidFill(sldFill, clrMapOvr, undefined, warpObj);
                bgcolor = `background: #${sldBgClr};`;
            }
            else if (bgFillTyp === "GRADIENT_FILL") {
                bgcolor = getBgGradientFill(bgPr, undefined, slideMasterContent, warpObj);
            }
            else if (bgFillTyp === "PIC_FILL") {
                bgcolor = await getBgPicFill(bgPr, "slideLayoutBg", warpObj, undefined);
            }
        }
        else if (bgRef !== undefined) {
            let phClr = getSolidFill(bgRef, clrMapOvr, undefined, warpObj);
            let idx = Number(bgRef["attrs"]["idx"]);
            if (idx == 0 || idx == 1000) ;
            else if (idx > 0 && idx < 1000) ;
            else if (idx > 1000) {
                let trueIdx = idx - 1000;
                let bgFillLst = warpObj["themeContent"]["a:theme"]["a:themeElements"]["a:fmtScheme"]["a:bgFillStyleLst"];
                let sortblAry = [];
                Object.keys(bgFillLst).forEach(key => {
                    let bgFillLstTyp = bgFillLst[key];
                    if (key != "attrs") {
                        if (bgFillLstTyp.constructor === Array) {
                            for (const item of bgFillLstTyp) {
                                let obj = {};
                                obj[key] = item;
                                obj["idex"] = item["attrs"]["order"];
                                obj["attrs"] = {
                                    "order": item["attrs"]["order"]
                                };
                                sortblAry.push(obj);
                            }
                        }
                        else {
                            let obj = {};
                            obj[key] = bgFillLstTyp;
                            obj["idex"] = bgFillLstTyp["attrs"]["order"];
                            obj["attrs"] = {
                                "order": bgFillLstTyp["attrs"]["order"]
                            };
                            sortblAry.push(obj);
                        }
                    }
                });
                let sortByOrder = sortblAry.slice(0);
                sortByOrder.sort((a, b) => {
                    return a.idex - b.idex;
                });
                let bgFillLstIdx = sortByOrder[trueIdx - 1];
                let bgFillTyp = getFillType(bgFillLstIdx);
                if (bgFillTyp === "SOLID_FILL") {
                    let sldFill = bgFillLstIdx["a:solidFill"];
                    let sldBgClr = getSolidFill(sldFill, clrMapOvr, phClr, warpObj);
                    bgcolor = `background: #${sldBgClr};`;
                }
                else if (bgFillTyp === "GRADIENT_FILL") {
                    bgcolor = getBgGradientFill(bgFillLstIdx, phClr, slideMasterContent, warpObj);
                }
                else if (bgFillTyp === "PIC_FILL") {
                    bgcolor = await getBgPicFill(bgFillLstIdx, "themeBg", warpObj, phClr);
                }
                else ;
            }
        }
        else {
            bgPr = PPTXXmlUtils.getTextByPathList(slideMasterContent, ["p:sldMaster", "p:cSld", "p:bg", "p:bgPr"]);
            bgRef = PPTXXmlUtils.getTextByPathList(slideMasterContent, ["p:sldMaster", "p:cSld", "p:bg", "p:bgRef"]);
            const clrMap = PPTXXmlUtils.getTextByPathList(slideMasterContent, ["p:sldMaster", "p:clrMap", "attrs"]);
            if (bgPr !== undefined) {
                let bgFillTyp = getFillType(bgPr);
                if (bgFillTyp === "SOLID_FILL") {
                    let sldFill = bgPr["a:solidFill"];
                    let sldBgClr = getSolidFill(sldFill, clrMap, undefined, warpObj);
                    bgcolor = `background: #${sldBgClr};`;
                }
                else if (bgFillTyp === "GRADIENT_FILL") {
                    bgcolor = getBgGradientFill(bgPr, undefined, slideMasterContent, warpObj);
                }
                else if (bgFillTyp === "PIC_FILL") {
                    bgcolor = await getBgPicFill(bgPr, "slideMasterBg", warpObj, undefined);
                }
            }
            else if (bgRef !== undefined) {
                let phClr = getSolidFill(bgRef, clrMap, undefined, warpObj);
                let idx = Number(bgRef["attrs"]["idx"]);
                if (idx == 0 || idx == 1000) ;
                else if (idx > 0 && idx < 1000) ;
                else if (idx > 1000) {
                    let trueIdx = idx - 1000;
                    let bgFillLst = warpObj["themeContent"]["a:theme"]["a:themeElements"]["a:fmtScheme"]["a:bgFillStyleLst"];
                    let sortblAry = [];
                    Object.keys(bgFillLst).forEach(key => {
                        let bgFillLstTyp = bgFillLst[key];
                        if (key != "attrs") {
                            if (bgFillLstTyp.constructor === Array) {
                                for (const item of bgFillLstTyp) {
                                    let obj = {};
                                    obj[key] = item;
                                    obj["idex"] = item["attrs"]["order"];
                                    obj["attrs"] = {
                                        "order": item["attrs"]["order"]
                                    };
                                    sortblAry.push(obj);
                                }
                            }
                            else {
                                let obj = {};
                                obj[key] = bgFillLstTyp;
                                obj["idex"] = bgFillLstTyp["attrs"]["order"];
                                obj["attrs"] = {
                                    "order": bgFillLstTyp["attrs"]["order"]
                                };
                                sortblAry.push(obj);
                            }
                        }
                    });
                    let sortByOrder = sortblAry.slice(0);
                    sortByOrder.sort((a, b) => {
                        return a.idex - b.idex;
                    });
                    let bgFillLstIdx = sortByOrder[trueIdx - 1];
                    let bgFillTyp = getFillType(bgFillLstIdx);
                    if (bgFillTyp == "SOLID_FILL") {
                        let sldFill = bgFillLstIdx["a:solidFill"];
                        let sldBgClr = getSolidFill(sldFill, clrMap, phClr, warpObj);
                        bgcolor = `background: #${sldBgClr}`;
                    }
                    else if (bgFillTyp == "GRADIENT_FILL") {
                        bgcolor = getBgGradientFill(bgFillLstIdx, phClr, slideMasterContent, warpObj);
                    }
                    else if (bgFillTyp == "PIC_FILL") {
                        bgcolor = await getBgPicFill(bgFillLstIdx, "themeBg", warpObj, phClr);
                    }
                    else ;
                }
            }
        }
    }
    return bgcolor;
}
function getBgGradientFill(bgPr, phClr, slideMasterContent, warpObj) {
    let bgcolor = "";
    if (bgPr !== undefined) {
        let grdFill = bgPr["a:gradFill"];
        let gsLst = grdFill["a:gsLst"]["a:gs"];
        if (!Array.isArray(gsLst)) {
            gsLst = gsLst ? [gsLst] : [];
        }
        let color_ary = [];
        const pos_ary = [];
        for (const i of gsLst.keys()) {
            let lo_color = getSolidFill(gsLst[i], slideMasterContent["p:sldMaster"]["p:clrMap"]["attrs"], phClr, warpObj);
            const pos = PPTXXmlUtils.getTextByPathList(gsLst[i], ["attrs", "pos"]);
            if (pos !== undefined) {
                pos_ary[i] = `${pos / 1000}%`;
            }
            else {
                pos_ary[i] = "";
            }
            color_ary[i] = `#${lo_color}`;
        }
        let lin = grdFill["a:lin"];
        let rot = 90;
        if (lin !== undefined) {
            rot = PPTXXmlUtils.angleToDegrees(lin["attrs"]["ang"]);
            rot = rot + 90;
        }
        bgcolor = `background: linear-gradient(${rot}deg,`;
        for (const i of gsLst.keys()) {
            if (i == gsLst.length - 1) {
                bgcolor += `${color_ary[i]} ${pos_ary[i]});`;
            }
            else {
                bgcolor += `${color_ary[i]} ${pos_ary[i]}, `;
            }
        }
    }
    else {
        if (phClr !== undefined) {
            bgcolor = `background: #${phClr};`;
        }
    }
    return bgcolor;
}
async function getBgPicFill(bgPr, sorce, warpObj, phClr, index) {
    let bgcolor;
    let picFillResult = await getPicFill(sorce, bgPr["a:blipFill"], warpObj);
    let picFillBase64 = picFillResult;
    if (typeof picFillResult === 'object' && picFillResult.img) {
        picFillBase64 = picFillResult.img;
    }
    let ordr = bgPr["attrs"]["order"];
    let aBlipNode = bgPr["a:blipFill"]["a:blip"];
    let duotone = PPTXXmlUtils.getTextByPathList(aBlipNode, ["a:duotone"]);
    if (duotone !== undefined) {
        let clr_ary = [];
        Object.keys(duotone).forEach(clr_type => {
            if (clr_type != "attrs") {
                let obj = {};
                obj[clr_type] = duotone[clr_type];
                clr_ary.push(getSolidFill(obj, undefined, phClr, warpObj));
            }
        });
    }
    let aphaModFixNode = PPTXXmlUtils.getTextByPathList(aBlipNode, ["a:alphaModFix", "attrs"]);
    let imgOpacity = "";
    if (aphaModFixNode !== undefined && aphaModFixNode["amt"] !== undefined && aphaModFixNode["amt"] != "") {
        const amt = parseInt(aphaModFixNode["amt"]) / 100000;
        imgOpacity = `opacity:${amt};`;
    }
    let prop_style = "";
    if (typeof picFillResult === 'object') {
        if (picFillResult.backgroundSize) {
            prop_style += `background-size: ${picFillResult.backgroundSize};`;
        }
        if (picFillResult.backgroundPosition) {
            prop_style += `background-position: ${picFillResult.backgroundPosition};`;
        }
        if (picFillResult.backgroundRepeat) {
            prop_style += `background-repeat: ${picFillResult.backgroundRepeat};`;
        }
    }
    bgcolor = `background: url(${picFillBase64});  z-index: ${ordr};${prop_style}${imgOpacity}`;
    return bgcolor;
}
function getGradientFill(node, warpObj) {
    let gsLst = node["a:gsLst"]["a:gs"];
    if (!Array.isArray(gsLst)) {
        gsLst = gsLst ? [gsLst] : [];
    }
    let color_ary = [];
    for (const i of gsLst.keys()) {
        let lo_color = getSolidFill(gsLst[i], undefined, undefined, warpObj);
        color_ary[i] = lo_color;
    }
    let lin = node["a:lin"];
    let rot = 0;
    if (lin !== undefined) {
        rot = PPTXXmlUtils.angleToDegrees(lin["attrs"]["ang"]) + 90;
    }
    return {
        "color": color_ary,
        "rot": rot
    };
}
async function getPicFill(type, node, warpObj) {
    let img;
    let rId = node["a:blip"]["attrs"]["r:embed"];
    let imgPath;
    if (type == "slideBg" || type == "slide") {
        imgPath = PPTXXmlUtils.getTextByPathList(warpObj, ["slideResObj", rId, "target"]);
    }
    else if (type == "slideLayoutBg") {
        imgPath = PPTXXmlUtils.getTextByPathList(warpObj, ["layoutResObj", rId, "target"]);
    }
    else if (type == "slideMasterBg") {
        imgPath = PPTXXmlUtils.getTextByPathList(warpObj, ["masterResObj", rId, "target"]);
    }
    else if (type == "themeBg") {
        imgPath = PPTXXmlUtils.getTextByPathList(warpObj, ["themeResObj", rId, "target"]);
    }
    else if (type == "diagramBg") {
        imgPath = PPTXXmlUtils.getTextByPathList(warpObj, ["diagramResObj", rId, "target"]);
    }
    if (imgPath === undefined) {
        return undefined;
    }
    img = PPTXXmlUtils.getTextByPathList(warpObj, ["loaded-images", imgPath]);
    if (img === undefined) {
        let context = 'slide';
        if (type == "slideMasterBg") {
            context = 'master';
        }
        else if (type == "slideLayoutBg") {
            context = 'layout';
        }
        imgPath = PPTXXmlUtils.resolveMediaPath(imgPath, context, '');
        let imgExt = imgPath.split(".").pop();
        if (imgExt == "xml") {
            return undefined;
        }
        let imgFile = warpObj["zip"].file(imgPath);
        if (imgFile === null || imgFile === undefined) {
            return undefined;
        }
        let imgArrayBuffer = await imgFile.async("arraybuffer");
        let imgMimeType = PPTXXmlUtils.getMimeType(imgExt);
        img = `data:${imgMimeType};base64,${PPTXXmlUtils.base64ArrayBuffer(imgArrayBuffer)}`;
        setTextByPathList(warpObj, ["loaded-images", imgPath], img);
    }
    let tileNode = node["a:tile"];
    let stretchNode = node["a:stretch"];
    let fillMode = "stretch";
    let backgroundSize = "cover";
    let backgroundPosition = "center";
    let backgroundRepeat = "no-repeat";
    if (tileNode) {
        fillMode = "tile";
        backgroundRepeat = "repeat";
        let sx = tileNode["attrs"]["sx"];
        let sy = tileNode["attrs"]["sy"];
        if (sx && sy) {
            let widthPercent = parseInt(sx) / 100000 * 100;
            let heightPercent = parseInt(sy) / 100000 * 100;
            backgroundSize = `${widthPercent}% ${heightPercent}%`;
        }
        let tx = tileNode["attrs"]["tx"];
        let ty = tileNode["attrs"]["ty"];
        if (tx && ty) {
            let xPercent = parseInt(tx) / 100000 * 100;
            let yPercent = parseInt(ty) / 100000 * 100;
            backgroundPosition = `${xPercent}% ${yPercent}%`;
        }
    }
    else if (stretchNode) {
        fillMode = "stretch";
        let fillRect = stretchNode["a:fillRect"];
        if (fillRect) {
            backgroundSize = "cover";
        }
    }
    return {
        "img": img,
        "fillMode": fillMode,
        "backgroundSize": backgroundSize,
        "backgroundPosition": backgroundPosition,
        "backgroundRepeat": backgroundRepeat
    };
}
function getPatternFill(node, warpObj) {
    let fgColor = "", bgColor = "", prst = "";
    let bgClr = node["a:bgClr"];
    let fgClr = node["a:fgClr"];
    prst = node["attrs"]["prst"];
    fgColor = getSolidFill(fgClr, undefined, undefined, warpObj);
    bgColor = getSolidFill(bgClr, undefined, undefined, warpObj);
    let linear_gradient = getLinerGrandient(prst, bgColor, fgColor);
    return linear_gradient;
}
function getLinerGrandient(prst, bgColor, fgColor) {
    switch (prst) {
        case "smGrid":
            return [`linear-gradient(to right,  #${fgColor} -1px, transparent 1px ), linear-gradient(to bottom,  #${fgColor} -1px, transparent 1px)  #${bgColor};`, "4px 4px"];
        case "dotGrid":
            return [`linear-gradient(to right,  #${fgColor} -1px, transparent 1px ), linear-gradient(to bottom,  #${fgColor} -1px, transparent 1px)  #${bgColor};`, "8px 8px"];
        case "lgGrid":
            return [`linear-gradient(to right,  #${fgColor} -1px, transparent 1.5px ), linear-gradient(to bottom,  #${fgColor} -1px, transparent 1.5px)  #${bgColor};`, "8px 8px"];
        case "wdUpDiag":
            return [`repeating-linear-gradient(-45deg, transparent 1px , transparent 4px, #${fgColor} 7px)#${bgColor};`];
        case "dkUpDiag":
            return [`repeating-linear-gradient(-45deg, transparent 1px , #${bgColor} 5px)#${fgColor};`];
        case "ltUpDiag":
            return [`repeating-linear-gradient(-45deg, transparent 1px , transparent 2px, #${fgColor} 4px)#${bgColor};`];
        case "wdDnDiag":
            return [`repeating-linear-gradient(45deg, transparent 1px , transparent 4px, #${fgColor} 7px)#${bgColor};`];
        case "dkDnDiag":
            return [`repeating-linear-gradient(45deg, transparent 1px , #${bgColor} 5px)#${fgColor};`];
        case "ltDnDiag":
            return [`repeating-linear-gradient(45deg, transparent 1px , transparent 2px, #${fgColor} 4px)#${bgColor};`];
        case "dkHorz":
            return [`repeating-linear-gradient(0deg, transparent 1px , transparent 2px, #${bgColor} 7px)#${fgColor};`];
        case "ltHorz":
            return [`repeating-linear-gradient(0deg, transparent 1px , transparent 5px, #${fgColor} 7px)#${bgColor};`];
        case "narHorz":
            return [`repeating-linear-gradient(0deg, transparent 1px , transparent 2px, #${fgColor} 4px)#${bgColor};`];
        case "dkVert":
            return [`repeating-linear-gradient(90deg, transparent 1px , transparent 2px, #${bgColor} 7px)#${fgColor};`];
        case "ltVert":
            return [`repeating-linear-gradient(90deg, transparent 1px , transparent 5px, #${fgColor} 7px)#${bgColor};`];
        case "narVert":
            return [`repeating-linear-gradient(90deg, transparent 1px , transparent 2px, #${fgColor} 4px)#${bgColor};`];
        case "lgCheck":
        case "smCheck":
            var size = "";
            let pos = "";
            if (prst == "lgCheck") {
                size = "8px 8px";
                pos = "0 0, 4px 4px, 4px 4px, 8px 8px";
            }
            else {
                size = "4px 4px";
                pos = "0 0, 2px 2px, 2px 2px, 4px 4px";
            }
            return [`linear-gradient(45deg,  #${fgColor} 25%, transparent 0, transparent 75%,  #${fgColor} 0), linear-gradient(45deg,  #${fgColor} 25%, transparent 0, transparent 75%,  #${fgColor} 0) #${bgColor};`, size, pos];
        case "dashUpDiag":
            return [`repeating-linear-gradient(152deg, #${fgColor}, #${fgColor} 5% , transparent 0, transparent 70%)#${bgColor};`, "4px 4px"];
        case "dashDnDiag":
            return [`repeating-linear-gradient(45deg, #${fgColor}, #${fgColor} 5% , transparent 0, transparent 70%)#${bgColor};`, "4px 4px"];
        case "diagBrick":
            return [`linear-gradient(45deg, transparent 15%,  #${fgColor} 30%, transparent 30%), linear-gradient(-45deg, transparent 15%,  #${fgColor} 30%, transparent 30%), linear-gradient(-45deg, transparent 65%,  #${fgColor} 80%, transparent 0) #${bgColor};`, "4px 4px"];
        case "horzBrick":
            return [`linear-gradient(335deg, #${bgColor} 1.6px, transparent 1.6px), linear-gradient(155deg, #${bgColor} 1.6px, transparent 1.6px), linear-gradient(335deg, #${bgColor} 1.6px, transparent 1.6px), linear-gradient(155deg, #${bgColor} 1.6px, transparent 1.6px) #${fgColor};`, "4px 4px", "0 0.15px, 0.3px 2.5px, 2px 2.15px, 2.35px 0.4px"];
        case "dashVert":
            return [`linear-gradient(0deg,  #${bgColor} 30%, transparent 30%),linear-gradient(90deg,transparent, transparent 40%, #${fgColor} 40%, #${fgColor} 60% , transparent 60%)#${bgColor};`, "4px 4px"];
        case "dashHorz":
            return [`linear-gradient(90deg,  #${bgColor} 30%, transparent 30%),linear-gradient(0deg,transparent, transparent 40%, #${fgColor} 40%, #${fgColor} 60% , transparent 60%)#${bgColor};`, "4px 4px"];
        case "solidDmnd":
            return [`linear-gradient(135deg,  #${fgColor} 25%, transparent 25%), linear-gradient(225deg,  #${fgColor} 25%, transparent 25%), linear-gradient(315deg,  #${fgColor} 25%, transparent 25%), linear-gradient(45deg,  #${fgColor} 25%, transparent 25%) #${bgColor};`, "8px 8px"];
        case "openDmnd":
            return [`linear-gradient(45deg, transparent 0%, transparent calc(50% - 0.5px),  #${fgColor} 50%, transparent calc(50% + 0.5px),  transparent 100%), linear-gradient(-45deg, transparent 0%, transparent calc(50% - 0.5px) , #${fgColor} 50%, transparent calc(50% + 0.5px),  transparent 100%) #${bgColor};`, "8px 8px"];
        case "dotDmnd":
            return [`radial-gradient(#${fgColor} 15%, transparent 0), radial-gradient(#${fgColor} 15%, transparent 0) #${bgColor};`, "4px 4px", "0 0, 2px 2px"];
        case "zigZag":
        case "wave":
            var size = "";
            if (prst == "zigZag")
                size = "0";
            else
                size = "1px";
            return [`linear-gradient(135deg,  #${fgColor} 25%, transparent 25%) 50px ${size}, linear-gradient(225deg,  #${fgColor} 25%, transparent 25%) 50px ${size}, linear-gradient(315deg,  #${fgColor} 25%, transparent 25%), linear-gradient(45deg,  #${fgColor} 25%, transparent 25%) #${bgColor};`, "4px 4px"];
        case "lgConfetti":
        case "smConfetti":
            var size = "";
            if (prst == "lgConfetti")
                size = "4px 4px";
            else
                size = "2px 2px";
            return [`linear-gradient(135deg,  #${fgColor} 25%, transparent 25%) 50px 1px, linear-gradient(225deg,  #${fgColor} 25%, transparent 25%), linear-gradient(315deg,  #${fgColor} 25%, transparent 25%) 50px 1px , linear-gradient(45deg,  #${fgColor} 25%, transparent 25%) #${bgColor};`, size];
        case "plaid":
            return [`linear-gradient(0deg, transparent, transparent 25%, #${fgColor}33 25%, #${fgColor}33 50%),linear-gradient(90deg, transparent, transparent 25%, #${fgColor}66 25%, #${fgColor}66 50%) #${bgColor};`, "4px 4px"];
        case "sphere":
            return [`radial-gradient(#${fgColor} 50%, transparent 50%),#${bgColor};`, "4px 4px"];
        case "weave":
        case "shingle":
            return [`linear-gradient(45deg, #${bgColor} 1.31px , #${fgColor} 1.4px, #${fgColor} 1.5px, transparent 1.5px, transparent 4.2px, #${fgColor} 4.2px, #${fgColor} 4.3px, transparent 4.31px), linear-gradient(-45deg,  #${bgColor} 1.31px , #${fgColor} 1.4px, #${fgColor} 1.5px, transparent 1.5px, transparent 4.2px, #${fgColor} 4.2px, #${fgColor} 4.3px, transparent 4.31px) 0 4px, #${bgColor};`, "4px 8px"];
        case "pct5":
        case "pct10":
        case "pct20":
        case "pct25":
        case "pct30":
        case "pct40":
        case "pct50":
        case "pct60":
        case "pct70":
        case "pct75":
        case "pct80":
        case "pct90":
        case "trellis":
        case "divot":
            let px_pr_ary;
            switch (prst) {
                case "pct5":
                    px_pr_ary = ["0.3px", "10%", "2px 2px"];
                    break;
                case "divot":
                    px_pr_ary = ["0.3px", "40%", "4px 4px"];
                    break;
                case "pct10":
                    px_pr_ary = ["0.3px", "20%", "2px 2px"];
                    break;
                case "pct20":
                    px_pr_ary = ["0.2px", "40%", "2px 2px"];
                    break;
                case "pct25":
                    px_pr_ary = ["0.2px", "50%", "2px 2px"];
                    break;
                case "pct30":
                    px_pr_ary = ["0.5px", "50%", "2px 2px"];
                    break;
                case "pct40":
                    px_pr_ary = ["0.5px", "70%", "2px 2px"];
                    break;
                case "pct50":
                    px_pr_ary = ["0.09px", "90%", "2px 2px"];
                    break;
                case "pct60":
                    px_pr_ary = ["0.3px", "90%", "2px 2px"];
                    break;
                case "pct70":
                case "trellis":
                    px_pr_ary = ["0.5px", "95%", "2px 2px"];
                    break;
                case "pct75":
                    px_pr_ary = ["0.65px", "100%", "2px 2px"];
                    break;
                case "pct80":
                    px_pr_ary = ["0.85px", "100%", "2px 2px"];
                    break;
                case "pct90":
                    px_pr_ary = ["1px", "100%", "2px 2px"];
                    break;
            }
            return [`radial-gradient(#${fgColor} ${px_pr_ary[0]}, transparent ${px_pr_ary[1]}),#${bgColor};`, px_pr_ary[2]];
        default:
            return [0, 0];
    }
}
function getSolidFill(node, clrMap, phClr, warpObj) {
    if (node === undefined) {
        return undefined;
    }
    let color = "";
    let clrNode;
    if (node["a:srgbClr"] !== undefined) {
        clrNode = node["a:srgbClr"];
        color = PPTXXmlUtils.getTextByPathList(clrNode, ["attrs", "val"]);
    }
    else if (node["a:schemeClr"] !== undefined) {
        clrNode = node["a:schemeClr"];
        let schemeClr = PPTXXmlUtils.getTextByPathList(clrNode, ["attrs", "val"]);
        color = getSchemeColorFromTheme(`a:${schemeClr}`, clrMap, phClr, warpObj);
    }
    else if (node["a:scrgbClr"] !== undefined) {
        clrNode = node["a:scrgbClr"];
        let defBultColorVals = clrNode["attrs"];
        let red = (defBultColorVals["r"].indexOf("%") != -1) ? defBultColorVals["r"].split("%").shift() : defBultColorVals["r"];
        let green = (defBultColorVals["g"].indexOf("%") != -1) ? defBultColorVals["g"].split("%").shift() : defBultColorVals["g"];
        let blue = (defBultColorVals["b"].indexOf("%") != -1) ? defBultColorVals["b"].split("%").shift() : defBultColorVals["b"];
        color = toHex(255 * (Number(red) / 100)) + toHex(255 * (Number(green) / 100)) + toHex(255 * (Number(blue) / 100));
    }
    else if (node["a:prstClr"] !== undefined) {
        clrNode = node["a:prstClr"];
        let prstClr = PPTXXmlUtils.getTextByPathList(clrNode, ["attrs", "val"]);
        color = getColorName2Hex(prstClr);
    }
    else if (node["a:hslClr"] !== undefined) {
        clrNode = node["a:hslClr"];
        let defBultColorVals = clrNode["attrs"];
        let hue = Number(defBultColorVals["hue"]) / 100000;
        let sat = Number((defBultColorVals["sat"].indexOf("%") != -1) ? defBultColorVals["sat"].split("%").shift() : defBultColorVals["sat"]) / 100;
        let lum = Number((defBultColorVals["lum"].indexOf("%") != -1) ? defBultColorVals["lum"].split("%").shift() : defBultColorVals["lum"]) / 100;
        let hsl2rgb = hslToRgb(hue, sat, lum);
        color = toHex(hsl2rgb.r) + toHex(hsl2rgb.g) + toHex(hsl2rgb.b);
    }
    else if (node["a:sysClr"] !== undefined) {
        clrNode = node["a:sysClr"];
        let sysClr = PPTXXmlUtils.getTextByPathList(clrNode, ["attrs", "lastClr"]);
        if (sysClr !== undefined) {
            color = sysClr;
        }
    }
    let isAlpha = false;
    let alpha = parseInt(PPTXXmlUtils.getTextByPathList(clrNode, ["a:alpha", "attrs", "val"])) / 100000;
    if (!isNaN(alpha)) {
        let al_color = tinycolor$1(color);
        al_color.setAlpha(alpha);
        color = al_color.toHex8();
        isAlpha = true;
    }
    let hueMod = parseInt(PPTXXmlUtils.getTextByPathList(clrNode, ["a:hueMod", "attrs", "val"])) / 100000;
    if (!isNaN(hueMod)) {
        color = applyHueMod(color, hueMod, isAlpha);
    }
    let lumMod = parseInt(PPTXXmlUtils.getTextByPathList(clrNode, ["a:lumMod", "attrs", "val"])) / 100000;
    if (!isNaN(lumMod)) {
        color = applyLumMod(color, lumMod, isAlpha);
    }
    let lumOff = parseInt(PPTXXmlUtils.getTextByPathList(clrNode, ["a:lumOff", "attrs", "val"])) / 100000;
    if (!isNaN(lumOff)) {
        color = applyLumOff(color, lumOff, isAlpha);
    }
    let satMod = parseInt(PPTXXmlUtils.getTextByPathList(clrNode, ["a:satMod", "attrs", "val"])) / 100000;
    if (!isNaN(satMod)) {
        color = applySatMod(color, satMod, isAlpha);
    }
    let shade = parseInt(PPTXXmlUtils.getTextByPathList(clrNode, ["a:shade", "attrs", "val"])) / 100000;
    if (!isNaN(shade)) {
        color = applyShade(color, shade, isAlpha);
    }
    let tint = parseInt(PPTXXmlUtils.getTextByPathList(clrNode, ["a:tint", "attrs", "val"])) / 100000;
    if (!isNaN(tint)) {
        color = applyTint(color, tint, isAlpha);
    }
    return color;
}
function toHex(n) {
    let hex = n.toString(16);
    while (hex.length < 2) {
        hex = `0${hex}`;
    }
    return hex;
}
function hslToRgb(hue, sat, light) {
    let t1, t2, r, g, b;
    hue = hue / 60;
    if (light <= 0.5) {
        t2 = light * (sat + 1);
    }
    else {
        t2 = light + sat - (light * sat);
    }
    t1 = light * 2 - t2;
    r = hueToRgb(t1, t2, hue + 2) * 255;
    g = hueToRgb(t1, t2, hue) * 255;
    b = hueToRgb(t1, t2, hue - 2) * 255;
    return { r: r, g: g, b: b };
}
function hueToRgb(t1, t2, hue) {
    if (hue < 0)
        hue += 6;
    if (hue >= 6)
        hue -= 6;
    if (hue < 1)
        return (t2 - t1) * hue + t1;
    else if (hue < 3)
        return t2;
    else if (hue < 4)
        return (t2 - t1) * (4 - hue) + t1;
    else
        return t1;
}
function getColorName2Hex(name) {
    let hex;
    let colorName = ['white', 'AliceBlue', 'AntiqueWhite', 'Aqua', 'Aquamarine', 'Azure', 'Beige', 'Bisque', 'black', 'BlanchedAlmond', 'Blue', 'BlueViolet', 'Brown', 'BurlyWood', 'CadetBlue', 'Chartreuse', 'Chocolate', 'Coral', 'CornflowerBlue', 'Cornsilk', 'Crimson', 'Cyan', 'DarkBlue', 'DarkCyan', 'DarkGoldenRod', 'DarkGray', 'DarkGrey', 'DarkGreen', 'DarkKhaki', 'DarkMagenta', 'DarkOliveGreen', 'DarkOrange', 'DarkOrchid', 'DarkRed', 'DarkSalmon', 'DarkSeaGreen', 'DarkSlateBlue', 'DarkSlateGray', 'DarkSlateGrey', 'DarkTurquoise', 'DarkViolet', 'DeepPink', 'DeepSkyBlue', 'DimGray', 'DimGrey', 'DodgerBlue', 'FireBrick', 'FloralWhite', 'ForestGreen', 'Fuchsia', 'Gainsboro', 'GhostWhite', 'Gold', 'GoldenRod', 'Gray', 'Grey', 'Green', 'GreenYellow', 'HoneyDew', 'HotPink', 'IndianRed', 'Indigo', 'Ivory', 'Khaki', 'Lavender', 'LavenderBlush', 'LawnGreen', 'LemonChiffon', 'LightBlue', 'LightCoral', 'LightCyan', 'LightGoldenRodYellow', 'LightGray', 'LightGrey', 'LightGreen', 'LightPink', 'LightSalmon', 'LightSeaGreen', 'LightSkyBlue', 'LightSlateGray', 'LightSlateGrey', 'LightSteelBlue', 'LightYellow', 'Lime', 'LimeGreen', 'Linen', 'Magenta', 'Maroon', 'MediumAquaMarine', 'MediumBlue', 'MediumOrchid', 'MediumPurple', 'MediumSeaGreen', 'MediumSlateBlue', 'MediumSpringGreen', 'MediumTurquoise', 'MediumVioletRed', 'MidnightBlue', 'MintCream', 'MistyRose', 'Moccasin', 'NavajoWhite', 'Navy', 'OldLace', 'Olive', 'OliveDrab', 'Orange', 'OrangeRed', 'Orchid', 'PaleGoldenRod', 'PaleGreen', 'PaleTurquoise', 'PaleVioletRed', 'PapayaWhip', 'PeachPuff', 'Peru', 'Pink', 'Plum', 'PowderBlue', 'Purple', 'RebeccaPurple', 'Red', 'RosyBrown', 'RoyalBlue', 'SaddleBrown', 'Salmon', 'SandyBrown', 'SeaGreen', 'SeaShell', 'Sienna', 'Silver', 'SkyBlue', 'SlateBlue', 'SlateGray', 'SlateGrey', 'Snow', 'SpringGreen', 'SteelBlue', 'Tan', 'Teal', 'Thistle', 'Tomato', 'Turquoise', 'Violet', 'Wheat', 'White', 'WhiteSmoke', 'Yellow', 'YellowGreen'];
    let colorHex = ['ffffff', 'f0f8ff', 'faebd7', '00ffff', '7fffd4', 'f0ffff', 'f5f5dc', 'ffe4c4', '000000', 'ffebcd', '0000ff', '8a2be2', 'a52a2a', 'deb887', '5f9ea0', '7fff00', 'd2691e', 'ff7f50', '6495ed', 'fff8dc', 'dc143c', '00ffff', '00008b', '008b8b', 'b8860b', 'a9a9a9', 'a9a9a9', '006400', 'bdb76b', '8b008b', '556b2f', 'ff8c00', '9932cc', '8b0000', 'e9967a', '8fbc8f', '483d8b', '2f4f4f', '2f4f4f', '00ced1', '9400d3', 'ff1493', '00bfff', '696969', '696969', '1e90ff', 'b22222', 'fffaf0', '228b22', 'ff00ff', 'dcdcdc', 'f8f8ff', 'ffd700', 'daa520', '808080', '808080', '008000', 'adff2f', 'f0fff0', 'ff69b4', 'cd5c5c', '4b0082', 'fffff0', 'f0e68c', 'e6e6fa', 'fff0f5', '7cfc00', 'fffacd', 'add8e6', 'f08080', 'e0ffff', 'fafad2', 'd3d3d3', 'd3d3d3', '90ee90', 'ffb6c1', 'ffa07a', '20b2aa', '87cefa', '778899', '778899', 'b0c4de', 'ffffe0', '00ff00', '32cd32', 'faf0e6', 'ff00ff', '800000', '66cdaa', '0000cd', 'ba55d3', '9370db', '3cb371', '7b68ee', '00fa9a', '48d1cc', 'c71585', '191970', 'f5fffa', 'ffe4e1', 'ffe4b5', 'ffdead', '000080', 'fdf5e6', '808000', '6b8e23', 'ffa500', 'ff4500', 'da70d6', 'eee8aa', '98fb98', 'afeeee', 'db7093', 'ffefd5', 'ffdab9', 'cd853f', 'ffc0cb', 'dda0dd', 'b0e0e6', '800080', '663399', 'ff0000', 'bc8f8f', '4169e1', '8b4513', 'fa8072', 'f4a460', '2e8b57', 'fff5ee', 'a0522d', 'c0c0c0', '87ceeb', '6a5acd', '708090', '708090', 'fffafa', '00ff7f', '4682b4', 'd2b48c', '008080', 'd8bfd8', 'ff6347', '40e0d0', 'ee82ee', 'f5deb3', 'ffffff', 'f5f5f5', 'ffff00', '9acd32'];
    let findIndx = colorName.indexOf(name);
    if (findIndx != -1) {
        hex = colorHex[findIndx];
    }
    return hex;
}
function getSchemeColorFromTheme(schemeClr, clrMap, phClr, warpObj) {
    let color = '';
    let slideLayoutClrOvride;
    if (clrMap !== undefined) {
        slideLayoutClrOvride = clrMap;
    }
    else if (warpObj !== undefined) {
        let sldClrMapOvr = PPTXXmlUtils.getTextByPathList(warpObj["slideContent"], ["p:sld", "p:clrMapOvr", "a:overrideClrMapping", "attrs"]);
        if (sldClrMapOvr !== undefined) {
            slideLayoutClrOvride = sldClrMapOvr;
        }
        else {
            let sldClrMapOvr = PPTXXmlUtils.getTextByPathList(warpObj["slideLayoutContent"], ["p:sldLayout", "p:clrMapOvr", "a:overrideClrMapping", "attrs"]);
            if (sldClrMapOvr !== undefined) {
                slideLayoutClrOvride = sldClrMapOvr;
            }
            else {
                slideLayoutClrOvride = PPTXXmlUtils.getTextByPathList(warpObj["slideMasterContent"], ["p:sldMaster", "p:clrMap", "attrs"]);
            }
        }
    }
    let schmClrName = schemeClr.substr(2);
    if (schmClrName == "phClr" && phClr !== undefined) {
        color = phClr;
    }
    else {
        if (slideLayoutClrOvride !== undefined) {
            switch (schmClrName) {
                case "tx1":
                case "tx2":
                case "bg1":
                case "bg2":
                    schemeClr = `a:${slideLayoutClrOvride[schmClrName]}`;
                    break;
            }
        }
        else {
            switch (schmClrName) {
                case "tx1":
                    schemeClr = "a:dk1";
                    break;
                case "tx2":
                    schemeClr = "a:dk2";
                    break;
                case "bg1":
                    schemeClr = "a:lt1";
                    break;
                case "bg2":
                    schemeClr = "a:lt2";
                    break;
            }
        }
        let refNode = PPTXXmlUtils.getTextByPathList(warpObj["themeContent"], ["a:theme", "a:themeElements", "a:clrScheme", schemeClr]);
        color = PPTXXmlUtils.getTextByPathList(refNode, ["a:srgbClr", "attrs", "val"]);
        if (color === undefined && refNode !== undefined) {
            color = PPTXXmlUtils.getTextByPathList(refNode, ["a:sysClr", "attrs", "lastClr"]);
        }
    }
    return color;
}
function extractChartData(serNode, warpObj) {
    let dataMat = new Array();
    if (serNode === undefined) {
        return dataMat;
    }
    if (serNode["c:xVal"] !== undefined) {
        var dataRow = new Array();
        eachElement(serNode["c:xVal"]["c:numRef"]["c:numCache"]["c:pt"], (innerNode, index) => {
            dataRow.push(parseFloat(innerNode["c:v"]));
            return "";
        });
        dataMat.push(dataRow);
        dataRow = new Array();
        eachElement(serNode["c:yVal"]["c:numRef"]["c:numCache"]["c:pt"], (innerNode, index) => {
            dataRow.push(parseFloat(innerNode["c:v"]));
            return "";
        });
        dataMat.push(dataRow);
    }
    else {
        eachElement(serNode, (innerNode, index) => {
            var dataRow = new Array();
            let colName;
            const txStrRef = PPTXXmlUtils.getTextByPathList(innerNode, ["c:tx", "c:strRef"]);
            if (txStrRef) {
                const strCache = PPTXXmlUtils.getTextByPathList(txStrRef, ["c:strCache"]);
                if (strCache) {
                    const pt = PPTXXmlUtils.getTextByPathList(strCache, ["c:pt"]);
                    if (pt) {
                        if (Array.isArray(pt)) {
                            colName = pt[0]["c:v"];
                        }
                        else {
                            colName = pt["c:v"];
                        }
                    }
                }
            }
            if (!colName) {
                colName = PPTXXmlUtils.getTextByPathList(innerNode, ["c:tx", "c:v"]) || index;
            }
            let rowNames = {};
            if (PPTXXmlUtils.getTextByPathList(innerNode, ["c:cat", "c:strRef", "c:strCache", "c:pt"]) !== undefined) {
                eachElement(innerNode["c:cat"]["c:strRef"]["c:strCache"]["c:pt"], (innerNode, index) => {
                    rowNames[innerNode["attrs"]["idx"]] = innerNode["c:v"];
                    return "";
                });
            }
            else if (PPTXXmlUtils.getTextByPathList(innerNode, ["c:cat", "c:numRef", "c:numCache", "c:pt"]) !== undefined) {
                eachElement(innerNode["c:cat"]["c:numRef"]["c:numCache"]["c:pt"], (innerNode, index) => {
                    rowNames[innerNode["attrs"]["idx"]] = innerNode["c:v"];
                    return "";
                });
            }
            else if (PPTXXmlUtils.getTextByPathList(innerNode, ["c:cat", "c:multiLvlStrRef", "c:multiLvlStrCache"]) !== undefined) {
                const multiLvlCache = PPTXXmlUtils.getTextByPathList(innerNode, ["c:cat", "c:multiLvlStrRef", "c:multiLvlStrCache"]);
                const lvl = PPTXXmlUtils.getTextByPathList(multiLvlCache, ["c:lvl"]);
                if (lvl) {
                    const firstLvl = Array.isArray(lvl) ? lvl[0] : lvl;
                    const pts = PPTXXmlUtils.getTextByPathList(firstLvl, ["c:pt"]);
                    if (pts) {
                        eachElement(pts, (pt, index) => {
                            rowNames[pt["attrs"]["idx"]] = pt["c:v"];
                            return "";
                        });
                    }
                }
            }
            if (PPTXXmlUtils.getTextByPathList(innerNode, ["c:val", "c:numRef", "c:numCache", "c:pt"]) !== undefined) {
                eachElement(innerNode["c:val"]["c:numRef"]["c:numCache"]["c:pt"], (innerNode, index) => {
                    dataRow.push({ x: innerNode["attrs"]["idx"], y: parseFloat(innerNode["c:v"]) });
                    return "";
                });
            }
            let seriesStyle = {};
            let fillType = getFillType(PPTXXmlUtils.getTextByPathList(innerNode, ["c:spPr"]));
            if (fillType === "SOLID_FILL" && warpObj !== undefined) {
                let fillNode = PPTXXmlUtils.getTextByPathList(innerNode, ["c:spPr", "a:solidFill"]);
                if (fillNode !== undefined) {
                    let fillColor = getSolidFill(fillNode, undefined, undefined, warpObj);
                    if (fillColor !== undefined) {
                        if (fillColor && !fillColor.startsWith('#')) {
                            fillColor = `#${fillColor}`;
                        }
                        seriesStyle.fillColor = fillColor;
                    }
                }
            }
            else if (fillType === "GRADIENT_FILL" && warpObj !== undefined) {
                let gradFillNode = PPTXXmlUtils.getTextByPathList(innerNode, ["c:spPr", "a:gradFill"]);
                if (gradFillNode !== undefined) {
                    let gradientFill = getGradientFill(gradFillNode, warpObj);
                    if (gradientFill !== undefined) {
                        seriesStyle.gradientFill = gradientFill;
                    }
                }
            }
            let lineNode = PPTXXmlUtils.getTextByPathList(innerNode, ["c:spPr", "a:ln"]);
            if (lineNode !== undefined && warpObj !== undefined) {
                let lineFillType = getFillType(lineNode);
                if (lineFillType === "SOLID_FILL") {
                    let lineColor = getSolidFill(lineNode["a:solidFill"], undefined, undefined, warpObj);
                    if (lineColor !== undefined) {
                        if (lineColor && !lineColor.startsWith('#')) {
                            lineColor = `#${lineColor}`;
                        }
                        seriesStyle.lineColor = lineColor;
                    }
                }
                else if (lineFillType === "GRADIENT_FILL") {
                    let lineGradFillNode = lineNode["a:gradFill"];
                    if (lineGradFillNode !== undefined) {
                        let lineGradientFill = getGradientFill(lineGradFillNode, warpObj);
                        if (lineGradientFill !== undefined) {
                            seriesStyle.lineGradientFill = lineGradientFill;
                        }
                    }
                }
            }
            dataMat.push({ key: colName, values: dataRow, xlabels: rowNames, style: seriesStyle });
            return "";
        });
    }
    return dataMat;
}
function setTextByPathList(node, path, value) {
    if (path.constructor !== Array) {
        throw Error("Error of path type! path is not array.");
    }
    if (node === undefined) {
        return undefined;
    }
    function setObjectPath(obj, parts, value) {
        if (!parts)
            return obj;
        let current = obj;
        let lent = parts.length;
        for (let i = 0; i < lent; i++) {
            const p = parts[i];
            if (current[p] === undefined) {
                if (i == lent - 1) {
                    current[p] = value;
                }
                else {
                    current[p] = {};
                }
            }
            current = current[p];
        }
        return obj;
    }
    setObjectPath(node, path, value);
}
function eachElement(node, doFunction) {
    if (node === undefined) {
        return;
    }
    let result = "";
    if (node.constructor === Array) {
        let l = node.length;
        for (let i = 0; i < l; i++) {
            result += doFunction(node[i], i);
        }
    }
    else {
        result += doFunction(node, 0);
    }
    return result;
}
function applyShade(rgbStr, shadeValue, isAlpha) {
    let color = tinycolor$1(rgbStr).toHsl();
    shadeValue = Math.max(0, Math.min(1, shadeValue));
    let cacl_l = Math.max(0, Math.min(1, color.l * shadeValue));
    if (isAlpha)
        return tinycolor$1({ h: color.h, s: color.s, l: cacl_l, a: color.a }).toHex8();
    return tinycolor$1({ h: color.h, s: color.s, l: cacl_l, a: color.a }).toHex();
}
function applyTint(rgbStr, tintValue, isAlpha) {
    let color = tinycolor$1(rgbStr).toHsl();
    tintValue = Math.max(0, Math.min(1, tintValue));
    let cacl_l = Math.max(0, Math.min(1, color.l * tintValue + (1 - tintValue)));
    if (isAlpha)
        return tinycolor$1({ h: color.h, s: color.s, l: cacl_l, a: color.a }).toHex8();
    return tinycolor$1({ h: color.h, s: color.s, l: cacl_l, a: color.a }).toHex();
}
function applyLumOff(rgbStr, offset, isAlpha) {
    let color = tinycolor$1(rgbStr).toHsl();
    let lum = offset + color.l;
    if (lum >= 1) {
        if (isAlpha)
            return tinycolor$1({ h: color.h, s: color.s, l: 1, a: color.a }).toHex8();
        return tinycolor$1({ h: color.h, s: color.s, l: 1, a: color.a }).toHex();
    }
    if (isAlpha)
        return tinycolor$1({ h: color.h, s: color.s, l: lum, a: color.a }).toHex8();
    return tinycolor$1({ h: color.h, s: color.s, l: lum, a: color.a }).toHex();
}
function applyLumMod(rgbStr, multiplier, isAlpha) {
    let color = tinycolor$1(rgbStr).toHsl();
    let cacl_l = color.l * multiplier;
    if (cacl_l >= 1) {
        cacl_l = 1;
    }
    if (isAlpha)
        return tinycolor$1({ h: color.h, s: color.s, l: cacl_l, a: color.a }).toHex8();
    return tinycolor$1({ h: color.h, s: color.s, l: cacl_l, a: color.a }).toHex();
}
function applyHueMod(rgbStr, multiplier, isAlpha) {
    let color = tinycolor$1(rgbStr).toHsl();
    let cacl_h = color.h * multiplier;
    if (cacl_h >= 360) {
        cacl_h = cacl_h - 360;
    }
    if (isAlpha)
        return tinycolor$1({ h: cacl_h, s: color.s, l: color.l, a: color.a }).toHex8();
    return tinycolor$1({ h: cacl_h, s: color.s, l: color.l, a: color.a }).toHex();
}
function applySatMod(rgbStr, multiplier, isAlpha) {
    let color = tinycolor$1(rgbStr).toHsl();
    let cacl_s = color.s * multiplier;
    if (cacl_s >= 1) {
        cacl_s = 1;
    }
    if (isAlpha)
        return tinycolor$1({ h: color.h, s: cacl_s, l: color.l, a: color.a }).toHex8();
    return tinycolor$1({ h: color.h, s: cacl_s, l: color.l, a: color.a }).toHex();
}
function rgba2hex(rgbaStr) {
    let a, rgb = rgbaStr.replace(/\s/g, '').match(/^rgba?\((\d+),(\d+),(\d+),?([^,\s)]+)?/i), alpha = (rgb && rgb[4] || "").trim(), hex = rgb ?
        (rgb[1] | 1 << 8).toString(16).slice(1) +
            (rgb[2] | 1 << 8).toString(16).slice(1) +
            (rgb[3] | 1 << 8).toString(16).slice(1) : rgbaStr;
    if (alpha !== "") {
        a = alpha;
    }
    else {
        a = 0o1;
    }
    a = ((a * 255) | 1 << 8).toString(16).slice(1);
    hex = hex + a;
    return hex;
}
function getSvgGradient(w, h, angl, color_arry, shpId) {
    const stopsArray = getMiddleStops(color_arry - 2);
    let svgAngle = '', svgHeight = h, svgWidth = w, svg = '', xy_ary = SVGangle(angl, svgHeight, svgWidth), x1 = xy_ary[0], y1 = xy_ary[1], x2 = xy_ary[2], y2 = xy_ary[3];
    let sal = stopsArray.length, sr = sal < 20 ? 100 : 1000;
    svgAngle = ` gradientUnits="userSpaceOnUse" x1="${x1}%" y1="${y1}%" x2="${x2}%" y2="${y2}%"`;
    svgAngle = `<linearGradient id="linGrd_${shpId}"${svgAngle}>\n`;
    svg += svgAngle;
    for (let i = 0; i < sal; i++) {
        const tinClr = tinycolor$1(`#${color_arry[i]}`);
        let alpha = tinClr.getAlpha();
        svg += `<stop offset="${Math.round(parseFloat(stopsArray[i]) / 100 * sr) / sr}" style="stop-color:${tinClr.toHexString()}; stop-opacity:${(alpha)};"`;
        svg += '/>\n';
    }
    svg += `</linearGradient>\n`;
    return svg;
}
function getMiddleStops(s) {
    let sArry = ['0%', '100%'];
    if (s == 0) {
        return sArry;
    }
    else {
        let i = s;
        while (i--) {
            let middleStop = 100 - ((100 / (s + 1)) * (i + 1)), middleStopString = `${middleStop}%`;
            sArry.splice(-1, 0, middleStopString);
        }
    }
    return sArry;
}
function SVGangle(deg, svgHeight, svgWidth) {
    let w = parseFloat(svgWidth), h = parseFloat(svgHeight), ang = parseFloat(deg), o = 2, n = 2, wc = w / 2, hc = h / 2, tx1 = 2, ty1 = 2, tx2 = 2, ty2 = 2, k = (((ang % 360) + 360) % 360), j = (360 - k) * Math.PI / 180, i = Math.tan(j), l = hc - i * wc;
    if (k == 0) {
        tx1 = w,
            ty1 = hc,
            tx2 = 0,
            ty2 = hc;
    }
    else if (k < 90) {
        n = w,
            o = 0;
    }
    else if (k == 90) {
        tx1 = wc,
            ty1 = 0,
            tx2 = wc,
            ty2 = h;
    }
    else if (k < 180) {
        n = 0,
            o = 0;
    }
    else if (k == 180) {
        tx1 = 0,
            ty1 = hc,
            tx2 = w,
            ty2 = hc;
    }
    else if (k < 270) {
        n = 0,
            o = h;
    }
    else if (k == 270) {
        tx1 = wc,
            ty1 = h,
            tx2 = wc,
            ty2 = 0;
    }
    else {
        n = w,
            o = h;
    }
    let m = o + (n / i);
    tx1 = tx1 == 2 ? i * (m - l) / (Math.pow(i, 2) + 1) : tx1,
        ty1 = ty1 == 2 ? i * tx1 + l : ty1,
        tx2 = tx2 == 2 ? w - tx1 : tx2,
        ty2 = ty2 == 2 ? h - ty1 : ty2;
    let x1 = Math.round(tx2 / w * 100 * 100) / 100, y1 = Math.round(ty2 / h * 100 * 100) / 100, x2 = Math.round(tx1 / w * 100 * 100) / 100, y2 = Math.round(ty1 / h * 100 * 100) / 100;
    return [x1, y1, x2, y2];
}
function getSvgImagePattern(node, fill, shpId, warpObj) {
    let fillUrl = fill;
    if (typeof fill === 'object' && fill.img) {
        fillUrl = fill.img;
    }
    let pic_dim = getBase64ImageDimensions(fillUrl);
    let width = pic_dim[0];
    let height = pic_dim[1];
    let blipFillNode = node["p:spPr"]["a:blipFill"];
    let sx = 0, sy = 0;
    let tileNode = PPTXXmlUtils.getTextByPathList(blipFillNode, ["a:tile", "attrs"]);
    if (tileNode !== undefined && tileNode["sx"] !== undefined) {
        sx = (parseInt(tileNode["sx"]) / 100000) * width;
        sy = (parseInt(tileNode["sy"]) / 100000) * height;
    }
    let blipNode = node["p:spPr"]["a:blipFill"]["a:blip"];
    let tialphaModFixNode = PPTXXmlUtils.getTextByPathList(blipNode, ["a:alphaModFix", "attrs"]);
    let imgOpacity = "";
    if (tialphaModFixNode !== undefined && tialphaModFixNode["amt"] !== undefined && tialphaModFixNode["amt"] != "") {
        parseInt(tialphaModFixNode["amt"]) / 100000;
    }
    let ptrn = '';
    if (sx !== undefined && sx != 0) {
        ptrn = `<pattern id="imgPtrn_${shpId}" x="0" y="0"  width="${sx}" height="${sy}" patternUnits="userSpaceOnUse">`;
    }
    else {
        ptrn = `<pattern id="imgPtrn_${shpId}"  patternContentUnits="objectBoundingBox"  width="1" height="1">`;
    }
    let duotoneNode = PPTXXmlUtils.getTextByPathList(blipNode, ["a:duotone"]);
    let fillterNode = "";
    let filterUrl = "";
    if (duotoneNode !== undefined) {
        const clr_ary = [];
        Object.keys(duotoneNode).forEach(clr_type => {
            if (clr_type != "attrs") {
                let obj = {};
                obj[clr_type] = duotoneNode[clr_type];
                let hexClr = getSolidFill(obj, undefined, undefined, warpObj);
                let color = tinycolor$1(`#${hexClr}`);
                clr_ary.push(color.toRgb());
            }
        });
        if (clr_ary.length == 2) {
            fillterNode = `<filter id="svg_image_duotone"> <feColorMatrix type="matrix" values=".33 .33 .33 0 0.33 .33 .33 0 0.33 .33 .33 0 00 0 0 1 0"></feColorMatrix><feComponentTransfer color-interpolation-filters="sRGB"><feFuncR type="table" tableValues="${clr_ary[0].r / 255} ${clr_ary[1].r / 255}"></feFuncR><feFuncG type="table" tableValues="${clr_ary[0].g / 255} ${clr_ary[1].g / 255}"></feFuncG><feFuncB type="table" tableValues="${clr_ary[0].b / 255} ${clr_ary[1].b / 255}"></feFuncB></feComponentTransfer> </filter>`;
        }
        filterUrl = 'filter="url(#svg_image_duotone)"';
        ptrn += fillterNode;
    }
    fillUrl = PPTXXmlUtils.escapeHtml(fillUrl);
    if (sx !== undefined && sx != 0) {
        ptrn += `<image  xlink:href="${fillUrl}" x="0" y="0" width="${sx}" height="${sy}" ${imgOpacity} ${filterUrl}></image>`;
    }
    else {
        ptrn += `<image  xlink:href="${fillUrl}" preserveAspectRatio="none" width="1" height="1" ${imgOpacity} ${filterUrl}></image>`;
    }
    ptrn += '</pattern>';
    return ptrn;
}
function getBase64ImageDimensions(imgSrc) {
    let image = new Image();
    image.onload = () => {
        image.width;
        image.height;
    };
    image.src = imgSrc;
    do {
        if (image.width !== undefined) {
            return [image.width, image.height];
        }
    } while (image.width === undefined);
}
function getVerticalAlign(node, slideLayoutSpNode, slideMasterSpNode, type) {
    let anchor = PPTXXmlUtils.getTextByPathList(node, ["p:txBody", "a:bodyPr", "attrs", "anchor"]);
    if (anchor === undefined) {
        anchor = PPTXXmlUtils.getTextByPathList(slideLayoutSpNode, ["p:txBody", "a:bodyPr", "attrs", "anchor"]);
        if (anchor === undefined) {
            anchor = PPTXXmlUtils.getTextByPathList(slideMasterSpNode, ["p:txBody", "a:bodyPr", "attrs", "anchor"]);
            if (anchor === undefined) {
                anchor = "t";
            }
        }
    }
    let shapeType = PPTXXmlUtils.getTextByPathList(node, ["p:spPr", "a:prstGeom", "attrs", "prst"]);
    const circularShapes = [
        "ellipse", "ovalCallout", "wedgeEllipseCallout",
        "pie", "pieWedge", "chord", "sector", "arc", "blockArc"
    ];
    if (circularShapes.includes(shapeType) && anchor === "t") {
        anchor = "ctr";
    }
    return (anchor === "ctr") ? "v-mid" : ((anchor === "b") ? "v-down" : "v-up");
}
function getContentDir(node, type, warpObj) {
    return "content";
}
function getVerticalMargins(pNode, textBodyNode, type, idx, warpObj, totalParagraphs, paragraphIndex, anchor) {
    let lvl = 1;
    let spcBefNode = PPTXXmlUtils.getTextByPathList(pNode, ["a:pPr", "a:spcBef", "a:spcPts", "attrs", "val"]);
    let spcAftNode = PPTXXmlUtils.getTextByPathList(pNode, ["a:pPr", "a:spcAft", "a:spcPts", "attrs", "val"]);
    let spcBefType = "Pts";
    let spcAftType = "Pts";
    const spcBefIsExplicit = (spcBefNode !== undefined);
    let spcBefScale = 1.0;
    if (!spcBefIsExplicit && totalParagraphs === 1 && paragraphIndex === 0) {
        spcBefScale = 0.0;
    }
    if (spcBefNode === undefined) {
        spcBefNode = PPTXXmlUtils.getTextByPathList(pNode, ["a:pPr", "a:spcBef", "a:spcPct", "attrs", "val"]);
        if (spcBefNode !== undefined) {
            spcBefType = "Pct";
        }
    }
    if (spcAftNode === undefined) {
        spcAftNode = PPTXXmlUtils.getTextByPathList(pNode, ["a:pPr", "a:spcAft", "a:spcPct", "attrs", "val"]);
        if (spcAftNode !== undefined) {
            spcAftType = "Pct";
        }
    }
    let lnSpcNode = PPTXXmlUtils.getTextByPathList(pNode, ["a:pPr", "a:lnSpc", "a:spcPct", "attrs", "val"]);
    let lnSpcNodeType = "Pct";
    if (lnSpcNode === undefined) {
        lnSpcNode = PPTXXmlUtils.getTextByPathList(pNode, ["a:pPr", "a:lnSpc", "a:spcPts", "attrs", "val"]);
        if (lnSpcNode !== undefined) {
            lnSpcNodeType = "Pts";
        }
    }
    let lvlNode = PPTXXmlUtils.getTextByPathList(pNode, ["a:pPr", "attrs", "lvl"]);
    if (lvlNode !== undefined) {
        lvl = parseInt(lvlNode) + 1;
    }
    let fontSize;
    if (PPTXXmlUtils.getTextByPathList(pNode, ["a:r"]) !== undefined) {
        let fontSizeStr = getFontSize(pNode["a:r"], textBodyNode, undefined, lvl, type, warpObj);
        if (fontSizeStr != "inherit") {
            const fontSizeMatch = fontSizeStr.match(/(\d+(?:\.\d+)?)px/);
            if (fontSizeMatch) {
                fontSize = parseFloat(fontSizeMatch[1]);
            }
        }
    }
    let lstStyle = textBodyNode["a:lstStyle"];
    if (lstStyle !== undefined && (spcBefNode === undefined || spcAftNode === undefined || lnSpcNode === undefined)) {
        let lvlKey = `a:lvl${lvl}pPr`;
        let lstLvlNode = lstStyle[lvlKey];
        if (lstLvlNode !== undefined) {
            if (spcBefNode === undefined) {
                spcBefNode = PPTXXmlUtils.getTextByPathList(lstLvlNode, ["a:spcBef", "a:spcPts", "attrs", "val"]);
                if (spcBefNode === undefined) {
                    spcBefNode = PPTXXmlUtils.getTextByPathList(lstLvlNode, ["a:spcBef", "a:spcPct", "attrs", "val"]);
                    if (spcBefNode !== undefined) {
                        spcBefType = "Pct";
                    }
                }
            }
            if (spcAftNode === undefined) {
                spcAftNode = PPTXXmlUtils.getTextByPathList(lstLvlNode, ["a:spcAft", "a:spcPts", "attrs", "val"]);
                if (spcAftNode === undefined) {
                    spcAftNode = PPTXXmlUtils.getTextByPathList(lstLvlNode, ["a:spcAft", "a:spcPct", "attrs", "val"]);
                    if (spcAftNode !== undefined) {
                        spcAftType = "Pct";
                    }
                }
            }
            if (lnSpcNode === undefined) {
                lnSpcNode = PPTXXmlUtils.getTextByPathList(lstLvlNode, ["a:lnSpc", "a:spcPct", "attrs", "val"]);
                if (lnSpcNode === undefined) {
                    lnSpcNode = PPTXXmlUtils.getTextByPathList(lstLvlNode, ["a:lnSpc", "a:spcPts", "attrs", "val"]);
                    if (lnSpcNode !== undefined) {
                        lnSpcNodeType = "Pts";
                    }
                }
            }
        }
    }
    let isInLayoutOrMaster = true;
    if (type == "shape" || type == "textBox") {
        isInLayoutOrMaster = false;
    }
    if (isInLayoutOrMaster && (spcBefNode === undefined || spcAftNode === undefined || lnSpcNode === undefined)) {
        if (idx !== undefined) {
            let laypPrNode = PPTXXmlUtils.getTextByPathList(warpObj, ["slideLayoutTables", "idxTable", idx, "p:txBody", "a:p", (lvl - 1), "a:pPr"]);
            if (spcBefNode === undefined) {
                spcBefNode = PPTXXmlUtils.getTextByPathList(laypPrNode, ["a:spcBef", "a:spcPts", "attrs", "val"]);
                if (spcBefNode === undefined) {
                    spcBefNode = PPTXXmlUtils.getTextByPathList(laypPrNode, ["a:spcBef", "a:spcPct", "attrs", "val"]);
                    if (spcBefNode !== undefined) {
                        spcBefType = "Pct";
                    }
                }
            }
            if (spcAftNode === undefined) {
                spcAftNode = PPTXXmlUtils.getTextByPathList(laypPrNode, ["a:spcAft", "a:spcPts", "attrs", "val"]);
                if (spcAftNode === undefined) {
                    spcAftNode = PPTXXmlUtils.getTextByPathList(laypPrNode, ["a:spcAft", "a:spcPct", "attrs", "val"]);
                    if (spcAftNode !== undefined) {
                        spcAftType = "Pct";
                    }
                }
            }
            if (lnSpcNode === undefined) {
                lnSpcNode = PPTXXmlUtils.getTextByPathList(laypPrNode, ["a:lnSpc", "a:spcPct", "attrs", "val"]);
                if (lnSpcNode === undefined) {
                    lnSpcNode = PPTXXmlUtils.getTextByPathList(laypPrNode, ["a:pPr", "a:lnSpc", "a:spcPts", "attrs", "val"]);
                    if (lnSpcNode !== undefined) {
                        lnSpcNodeType = "Pts";
                    }
                }
            }
        }
    }
    if (isInLayoutOrMaster && (spcBefNode === undefined || spcAftNode === undefined || lnSpcNode === undefined)) {
        const slideMasterTextStyles = warpObj["slideMasterTextStyles"];
        let dirLoc = "";
        lvl = `a:lvl${lvl}pPr`;
        switch (type) {
            case "title":
            case "ctrTitle":
                dirLoc = "p:titleStyle";
                break;
            case "body":
            case "obj":
            case "dt":
            case "ftr":
            case "sldNum":
            case "textBox":
                dirLoc = "p:bodyStyle";
                break;
            case "shape":
            default:
                dirLoc = "p:otherStyle";
        }
        let inLvlNode = PPTXXmlUtils.getTextByPathList(slideMasterTextStyles, [dirLoc, lvl]);
        if (inLvlNode !== undefined) {
            if (spcBefNode === undefined) {
                spcBefNode = PPTXXmlUtils.getTextByPathList(inLvlNode, ["a:spcBef", "a:spcPts", "attrs", "val"]);
                if (spcBefNode === undefined) {
                    spcBefNode = PPTXXmlUtils.getTextByPathList(inLvlNode, ["a:spcBef", "a:spcPct", "attrs", "val"]);
                    if (spcBefNode !== undefined) {
                        spcBefType = "Pct";
                    }
                }
            }
            if (spcAftNode === undefined) {
                spcAftNode = PPTXXmlUtils.getTextByPathList(inLvlNode, ["a:spcAft", "a:spcPts", "attrs", "val"]);
                if (spcAftNode === undefined) {
                    spcAftNode = PPTXXmlUtils.getTextByPathList(inLvlNode, ["a:spcAft", "a:spcPct", "attrs", "val"]);
                    if (spcAftNode !== undefined) {
                        spcAftType = "Pct";
                    }
                }
            }
            if (lnSpcNode === undefined) {
                lnSpcNode = PPTXXmlUtils.getTextByPathList(inLvlNode, ["a:lnSpc", "a:spcPct", "attrs", "val"]);
                if (lnSpcNode === undefined) {
                    lnSpcNode = PPTXXmlUtils.getTextByPathList(inLvlNode, ["a:pPr", "a:lnSpc", "a:spcPts", "attrs", "val"]);
                    if (lnSpcNode !== undefined) {
                        lnSpcNodeType = "Pts";
                    }
                }
            }
        }
    }
    let spcBefor = 0, spcAfter = 0, spcLines = 0;
    let marginTopBottomStr = "";
    if (spcBefNode !== undefined) {
        if (spcBefType === "Pct") {
            spcBefor = parseInt(spcBefNode) / 1000;
        }
        else {
            spcBefor = parseInt(spcBefNode) / 100;
        }
    }
    if (spcAftNode !== undefined) {
        if (spcAftType === "Pct") {
            spcAfter = parseInt(spcAftNode) / 1000;
        }
        else {
            spcAfter = parseInt(spcAftNode) / 100;
        }
    }
    if (lnSpcNode !== undefined) {
        if (lnSpcNodeType === "Pct") {
            spcLines = parseInt(lnSpcNode) / 100000;
            let lineHeight = spcLines;
            if (lineHeight < 1.0) {
                lineHeight = 1.3;
            }
            marginTopBottomStr += `line-height: ${lineHeight};`;
        }
        else if (lnSpcNodeType === "Pts") {
            spcLines = parseInt(lnSpcNode) / 100;
            if (fontSize && fontSize > 0) {
                let lineHeight = spcLines / fontSize;
                if (lineHeight < 1.0) {
                    lineHeight = 1.3;
                }
                marginTopBottomStr += `line-height: ${lineHeight};`;
            }
        }
    }
    else if (type === "textBox") {
        marginTopBottomStr += "line-height: 1.3;";
    }
    if (spcBefNode !== undefined && (spcBefIsExplicit || spcBefScale > 0) && anchor !== "ctr") {
        let marginTop;
        if (spcBefType === "Pct") {
            let lineHeightPx = fontSize || 18;
            let spcBeforLimited = Math.min(spcBefor, 0.5);
            spcBeforLimited *= spcBefScale;
            marginTop = lineHeightPx * spcBeforLimited;
        }
        else {
            marginTop = spcBefor * 1.33 * spcBefScale;
        }
        if (marginTop > 0) {
            marginTop = Math.round(marginTop * 100) / 100;
            marginTopBottomStr += `margin-top: ${marginTop}px;`;
        }
    }
    if (spcAftNode !== undefined && anchor !== "ctr") {
        let marginBottom;
        if (spcAftType === "Pct") {
            let lineHeightPx = fontSize || 18;
            let spcAfterLimited = Math.min(spcAfter, 0.5);
            marginBottom = lineHeightPx * spcAfterLimited;
        }
        else {
            marginBottom = spcAfter * 1.33;
        }
        marginBottom = Math.round(marginBottom * 100) / 100;
        marginTopBottomStr += `margin-bottom: ${marginBottom}px;`;
    }
    return marginTopBottomStr;
}
function getHorizontalAlign(node, textBodyNode, idx, type, prg_dir, warpObj, spNode) {
    let algn = PPTXXmlUtils.getTextByPathList(node, ["a:pPr", "attrs", "algn"]);
    if (algn === undefined) {
        let layoutMasterNode = getLayoutAndMasterNode(node, idx, type, warpObj);
        let { nodeLaout: pPrNodeLaout, nodeMaster: pPrNodeMaster } = layoutMasterNode;
        let lvlIdx = 1;
        let lvlNode = PPTXXmlUtils.getTextByPathList(node, ["a:pPr", "attrs", "lvl"]);
        if (lvlNode !== undefined) {
            lvlIdx = parseInt(lvlNode) + 1;
        }
        let lvlStr = `a:lvl${lvlIdx}pPr`;
        let lstStyle = textBodyNode["a:lstStyle"];
        algn = PPTXXmlUtils.getTextByPathList(lstStyle, [lvlStr, "attrs", "algn"]);
        if (algn === undefined && idx !== undefined) {
            algn = PPTXXmlUtils.getTextByPathList(warpObj["slideLayoutTables"]["idxTable"][idx], ["p:txBody", "a:lstStyle", lvlStr, "attrs", "algn"]);
            if (algn === undefined) {
                algn = PPTXXmlUtils.getTextByPathList(warpObj["slideLayoutTables"]["idxTable"][idx], ["p:txBody", "a:p", "a:pPr", "attrs", "algn"]);
                if (algn === undefined) {
                    algn = PPTXXmlUtils.getTextByPathList(warpObj["slideLayoutTables"]["idxTable"][idx], ["p:txBody", "a:p", (lvlIdx - 1), "a:pPr", "attrs", "algn"]);
                }
            }
        }
        if (algn === undefined) {
            if (type !== undefined) {
                algn = PPTXXmlUtils.getTextByPathList(warpObj, ["slideLayoutTables", "typeTable", type, "p:txBody", "a:lstStyle", lvlStr, "attrs", "algn"]);
                if (algn === undefined) {
                    if (type == "title" || type == "ctrTitle") {
                        algn = PPTXXmlUtils.getTextByPathList(warpObj, ["slideMasterTextStyles", "p:titleStyle", lvlStr, "attrs", "algn"]);
                    }
                    else if (type == "body" || type == "obj" || type == "subTitle") {
                        algn = PPTXXmlUtils.getTextByPathList(warpObj, ["slideMasterTextStyles", "p:bodyStyle", lvlStr, "attrs", "algn"]);
                    }
                    else if (type == "shape" || type == "diagram") {
                        algn = PPTXXmlUtils.getTextByPathList(warpObj, ["slideMasterTextStyles", "p:otherStyle", lvlStr, "attrs", "algn"]);
                    }
                    else if (type == "textBox") {
                        algn = PPTXXmlUtils.getTextByPathList(warpObj, ["defaultTextStyle", lvlStr, "attrs", "algn"]);
                    }
                    else {
                        algn = PPTXXmlUtils.getTextByPathList(warpObj, ["slideMasterTables", "typeTable", type, "p:txBody", "a:lstStyle", lvlStr, "attrs", "algn"]);
                    }
                }
            }
            else {
                algn = PPTXXmlUtils.getTextByPathList(warpObj, ["slideMasterTextStyles", "p:bodyStyle", lvlStr, "attrs", "algn"]);
            }
        }
        if (algn === undefined && pPrNodeLaout) {
            algn = PPTXXmlUtils.getTextByPathList(pPrNodeLaout, ["attrs", "algn"]);
        }
        if (algn === undefined && pPrNodeMaster) {
            algn = PPTXXmlUtils.getTextByPathList(pPrNodeMaster, ["attrs", "algn"]);
        }
    }
    if (algn === undefined) {
        if (type == "title" || type == "subTitle" || type == "ctrTitle") {
            return "h-mid";
        }
        else if (type == "sldNum") {
            return "h-right";
        }
        else {
            return "h-left";
        }
    }
    let shapeType = "";
    if (spNode) {
        shapeType = PPTXXmlUtils.getTextByPathList(spNode, ["p:spPr", "a:prstGeom", "attrs", "prst"]);
    }
    if (!shapeType && spNode && spNode["attrs"] && spNode["attrs"]["data-geom-type"]) {
        shapeType = spNode["attrs"]["data-geom-type"];
    }
    const circularShapes = [
        "ellipse", "ovalCallout", "wedgeEllipseCallout",
        "pie", "pieWedge", "chord", "sector", "arc", "blockArc"
    ];
    const isCircularShape = circularShapes.includes(shapeType);
    if (isCircularShape) {
        return "h-mid";
    }
    if (algn !== undefined) {
        switch (algn) {
            case "l":
                if (prg_dir == "pregraph-rtl") {
                    return "h-left-rtl";
                }
                else {
                    return "h-left";
                }
            case "r":
                if (prg_dir == "pregraph-rtl") {
                    return "h-right-rtl";
                }
                else {
                    return "h-right";
                }
            case "ctr":
                return "h-mid";
            case "just":
            case "dist":
            default:
                return `h-${algn}`;
        }
    }
}
function getLayoutAndMasterNode(node, idx, type, warpObj) {
    let pPrNodeLaout, pPrNodeMaster;
    const pPrNode = node["a:pPr"];
    let lvl = 1;
    let lvlNode = PPTXXmlUtils.getTextByPathList(pPrNode, ["attrs", "lvl"]);
    if (lvlNode !== undefined) {
        lvl = parseInt(lvlNode) + 1;
    }
    if (idx !== undefined) {
        pPrNodeLaout = PPTXXmlUtils.getTextByPathList(warpObj["slideLayoutTables"]["idxTable"][idx], ["p:txBody", "a:lstStyle", `a:lvl${lvl}pPr`]);
        if (pPrNodeLaout === undefined) {
            pPrNodeLaout = PPTXXmlUtils.getTextByPathList(warpObj["slideLayoutTables"]["idxTable"][idx], ["p:txBody", "a:p", "a:pPr"]);
            if (pPrNodeLaout === undefined) {
                pPrNodeLaout = PPTXXmlUtils.getTextByPathList(warpObj["slideLayoutTables"]["idxTable"][idx], ["p:txBody", "a:p", (lvl - 1), "a:pPr"]);
            }
        }
    }
    if (type !== undefined) {
        let lvlStr = `a:lvl${lvl}pPr`;
        if (pPrNodeLaout === undefined) {
            pPrNodeLaout = PPTXXmlUtils.getTextByPathList(warpObj, ["slideLayoutTables", "typeTable", type, "p:txBody", "a:lstStyle", lvlStr]);
        }
        if (type == "title" || type == "ctrTitle") {
            pPrNodeMaster = PPTXXmlUtils.getTextByPathList(warpObj, ["slideMasterTextStyles", "p:titleStyle", lvlStr]);
        }
        else if (type == "body" || type == "obj" || type == "subTitle") {
            pPrNodeMaster = PPTXXmlUtils.getTextByPathList(warpObj, ["slideMasterTextStyles", "p:bodyStyle", lvlStr]);
        }
        else if (type == "shape" || type == "diagram") {
            pPrNodeMaster = PPTXXmlUtils.getTextByPathList(warpObj, ["slideMasterTextStyles", "p:otherStyle", lvlStr]);
        }
        else if (type == "textBox") {
            pPrNodeMaster = PPTXXmlUtils.getTextByPathList(warpObj, ["defaultTextStyle", lvlStr]);
        }
        else {
            pPrNodeMaster = PPTXXmlUtils.getTextByPathList(warpObj, ["slideMasterTables", "typeTable", type, "p:txBody", "a:lstStyle", lvlStr]);
        }
    }
    return {
        "nodeLaout": pPrNodeLaout,
        "nodeMaster": pPrNodeMaster
    };
}
function getPregraphDir(node, textBodyNode, idx, type, warpObj) {
    let rtl = PPTXXmlUtils.getTextByPathList(node, ["a:pPr", "attrs", "rtl"]);
    if (rtl === undefined) {
        let layoutMasterNode = getLayoutAndMasterNode(node, idx, type, warpObj);
        let { nodeLaout: pPrNodeLaout, nodeMaster: pPrNodeMaster } = layoutMasterNode;
        rtl = PPTXXmlUtils.getTextByPathList(pPrNodeLaout, ["attrs", "rtl"]);
        if (rtl === undefined && type != "shape") {
            rtl = PPTXXmlUtils.getTextByPathList(pPrNodeMaster, ["attrs", "rtl"]);
        }
    }
    if (rtl === undefined && textBodyNode !== undefined) {
        let rtlCol = PPTXXmlUtils.getTextByPathList(textBodyNode, ["a:bodyPr", "attrs", "rtlCol"]);
        if (rtlCol !== undefined) {
            rtl = rtlCol;
        }
    }
    if (rtl == "1") {
        return "pregraph-rtl";
    }
    else if (rtl == "0") {
        return "pregraph-ltr";
    }
    return "pregraph-inherit";
}
function getPregraphMargn(pNode, idx, type, isBullate, warpObj, fontSize) {
    if (!isBullate) {
        return ["", 0];
    }
    let marLStr = "", maginVal = 0;
    let pPrNode = pNode["a:pPr"];
    let layoutMasterNode = getLayoutAndMasterNode(pNode, idx, type, warpObj);
    let { nodeLaout: pPrNodeLaout, nodeMaster: pPrNodeMaster } = layoutMasterNode;
    let getRtlVal = PPTXXmlUtils.getTextByPathList(pPrNode, ["attrs", "rtl"]);
    if (getRtlVal === undefined) {
        getRtlVal = PPTXXmlUtils.getTextByPathList(pPrNodeLaout, ["attrs", "rtl"]);
        if (getRtlVal === undefined && type != "shape") {
            getRtlVal = PPTXXmlUtils.getTextByPathList(pPrNodeMaster, ["attrs", "rtl"]);
        }
    }
    let alignNode = PPTXXmlUtils.getTextByPathList(pPrNode, ["attrs", "algn"]);
    if (alignNode === undefined) {
        alignNode = PPTXXmlUtils.getTextByPathList(pPrNodeLaout, ["attrs", "algn"]);
        if (alignNode === undefined) {
            alignNode = PPTXXmlUtils.getTextByPathList(pPrNodeMaster, ["attrs", "algn"]);
        }
    }
    let indentNode = PPTXXmlUtils.getTextByPathList(pPrNode, ["attrs", "indent"]);
    if (indentNode === undefined) {
        indentNode = PPTXXmlUtils.getTextByPathList(pPrNodeLaout, ["attrs", "indent"]);
        if (indentNode === undefined) {
            indentNode = PPTXXmlUtils.getTextByPathList(pPrNodeMaster, ["attrs", "indent"]);
        }
    }
    let indent = 0;
    if (indentNode !== undefined) {
        indent = parseInt(indentNode) * SLIDE_FACTOR$1;
    }
    let marLNode = PPTXXmlUtils.getTextByPathList(pPrNode, ["attrs", "marL"]);
    if (marLNode === undefined) {
        marLNode = PPTXXmlUtils.getTextByPathList(pPrNodeLaout, ["attrs", "marL"]);
        if (marLNode === undefined) {
            marLNode = PPTXXmlUtils.getTextByPathList(pPrNodeMaster, ["attrs", "marL"]);
        }
    }
    let marginLeft = 0;
    if (marLNode !== undefined) {
        marginLeft = parseInt(marLNode) * SLIDE_FACTOR$1;
    }
    if ((indentNode !== undefined || marLNode !== undefined)) {
        marLStr = "padding-left: ";
        if (isBullate) {
            maginVal = Math.abs(0 - indent);
            let bulletSizeAdjustment = 0;
            if (fontSize !== undefined) {
                bulletSizeAdjustment = fontSize * 0.8;
            }
            maginVal = Math.max(0, maginVal - bulletSizeAdjustment);
            marLStr += `${maginVal}px;`;
        }
        else {
            maginVal = Math.abs(marginLeft + indent);
            marLStr += `${maginVal}px;`;
        }
    }
    let marRNode = PPTXXmlUtils.getTextByPathList(pPrNode, ["attrs", "marR"]);
    if (marRNode === undefined && marLNode === undefined) {
        marRNode = PPTXXmlUtils.getTextByPathList(pPrNodeLaout, ["attrs", "marR"]);
        if (marRNode === undefined) {
            marRNode = PPTXXmlUtils.getTextByPathList(pPrNodeMaster, ["attrs", "marR"]);
        }
    }
    return [marLStr, maginVal];
}
function extractChartTitleStyle(chartNode, warpObj) {
    const titleNode = PPTXXmlUtils.getTextByPathList(chartNode, ["c:title"]);
    if (!titleNode)
        return { text: "", style: {} };
    const style = {};
    let text = "";
    const rich = PPTXXmlUtils.getTextByPathList(titleNode, ["c:tx", "c:rich"]);
    if (rich) {
        const p = PPTXXmlUtils.getTextByPathList(rich, ["a:p"]);
        if (p) {
            const r = PPTXXmlUtils.getTextByPathList(p, ["a:r"]);
            if (r) {
                if (Array.isArray(r)) {
                    const textArray = r.map(run => PPTXXmlUtils.getTextByPathList(run, ["a:t"]));
                    text = textArray.filter(t => t).join('');
                }
                else {
                    text = PPTXXmlUtils.getTextByPathList(r, ["a:t"]);
                }
            }
            const pPr = PPTXXmlUtils.getTextByPathList(p, ["a:pPr"]);
            if (pPr) {
                const defRPr = PPTXXmlUtils.getTextByPathList(pPr, ["a:defRPr"]);
                if (defRPr) {
                    if (defRPr["attrs"] && defRPr["attrs"]["sz"]) {
                        style.fontSize = parseFloat(defRPr["attrs"]["sz"]) / 100;
                    }
                    if (defRPr["attrs"] && defRPr["attrs"]["b"] === "1") {
                        style.fontWeight = "bold";
                    }
                    const solidFill = PPTXXmlUtils.getTextByPathList(defRPr, ["a:solidFill"]);
                    if (solidFill) {
                        let color = getColor(solidFill, undefined, undefined, warpObj);
                        if (color && !color.startsWith('#')) {
                            color = `#${color}`;
                        }
                        style.color = color;
                    }
                }
            }
        }
    }
    if (!text) {
        const txPr = PPTXXmlUtils.getTextByPathList(titleNode, ["c:txPr"]);
        if (txPr) {
            const p = PPTXXmlUtils.getTextByPathList(txPr, ["a:p"]);
            if (p) {
                const r = PPTXXmlUtils.getTextByPathList(p, ["a:r"]);
                if (r) {
                    if (Array.isArray(r)) {
                        const textArray = r.map(run => PPTXXmlUtils.getTextByPathList(run, ["a:t"]));
                        text = textArray.filter(t => t).join('');
                    }
                    else {
                        text = PPTXXmlUtils.getTextByPathList(r, ["a:t"]);
                    }
                }
                const pPr = PPTXXmlUtils.getTextByPathList(p, ["a:pPr"]);
                if (pPr) {
                    const defRPr = PPTXXmlUtils.getTextByPathList(pPr, ["a:defRPr"]);
                    if (defRPr) {
                        if (defRPr["attrs"] && defRPr["attrs"]["sz"]) {
                            style.fontSize = parseFloat(defRPr["attrs"]["sz"]) / 100;
                        }
                        if (defRPr["attrs"] && defRPr["attrs"]["b"] === "1") {
                            style.fontWeight = "bold";
                        }
                        const solidFill = PPTXXmlUtils.getTextByPathList(defRPr, ["a:solidFill"]);
                        if (solidFill) {
                            let color = getColor(solidFill, undefined, undefined, warpObj);
                            if (color && !color.startsWith('#')) {
                                color = `#${color}`;
                            }
                            style.color = color;
                        }
                    }
                }
            }
        }
    }
    if (!text) {
        const tx = PPTXXmlUtils.getTextByPathList(titleNode, ["c:tx", "c:strRef", "c:strCache", "c:pt", "c:v"]);
        if (tx) {
            text = tx;
        }
    }
    return { text, style };
}
function extractChartAreaStyle(chartSpaceNode, warpObj) {
    const style = {};
    const spPr = PPTXXmlUtils.getTextByPathList(chartSpaceNode, ["c:spPr"]);
    if (spPr) {
        const fillType = getFillType(spPr);
        if (fillType === "SOLID_FILL") {
            const solidFill = PPTXXmlUtils.getTextByPathList(spPr, ["a:solidFill"]);
            if (solidFill) {
                let fillColor = getSolidFill(solidFill, undefined, undefined, warpObj);
                if (fillColor && !fillColor.startsWith('#')) {
                    fillColor = `#${fillColor}`;
                }
                style.fillColor = fillColor;
            }
        }
        else if (fillType === "GRADIENT_FILL") {
            const gradFill = PPTXXmlUtils.getTextByPathList(spPr, ["a:gradFill"]);
            if (gradFill) {
                style.gradientFill = getGradientFill(gradFill, warpObj);
            }
        }
        const ln = PPTXXmlUtils.getTextByPathList(spPr, ["a:ln"]);
        if (ln) {
            const solidFill = PPTXXmlUtils.getTextByPathList(ln, ["a:solidFill"]);
            if (solidFill) {
                let borderColor = getSolidFill(solidFill, undefined, undefined, warpObj);
                if (borderColor && !borderColor.startsWith('#')) {
                    borderColor = `#${borderColor}`;
                }
                style.borderColor = borderColor;
            }
            if (ln["attrs"] && ln["attrs"]["w"]) {
                style.borderWidth = parseFloat(ln["attrs"]["w"]) / 9525;
            }
        }
    }
    return style;
}
function extractChartLegendStyle(chartNode, warpObj) {
    const legendNode = PPTXXmlUtils.getTextByPathList(chartNode, ["c:legend"]);
    if (!legendNode)
        return {};
    const style = {};
    if (legendNode["c:legendPos"]) {
        style.position = legendNode["c:legendPos"]["attrs"]["val"];
    }
    const txPr = PPTXXmlUtils.getTextByPathList(legendNode, ["c:txPr"]);
    if (txPr) {
        const p = PPTXXmlUtils.getTextByPathList(txPr, ["a:p"]);
        if (p) {
            const pPr = PPTXXmlUtils.getTextByPathList(p, ["a:pPr"]);
            if (pPr) {
                const defRPr = PPTXXmlUtils.getTextByPathList(pPr, ["a:defRPr"]);
                if (defRPr) {
                    if (defRPr["attrs"] && defRPr["attrs"]["sz"]) {
                        style.fontSize = parseFloat(defRPr["attrs"]["sz"]) / 100;
                    }
                    const solidFill = PPTXXmlUtils.getTextByPathList(defRPr, ["a:solidFill"]);
                    if (solidFill) {
                        let color = getSolidFill(solidFill, undefined, undefined, warpObj);
                        if (color && !color.startsWith('#')) {
                            color = `#${color}`;
                        }
                        style.color = color;
                    }
                }
            }
        }
    }
    return style;
}
function extractChartAxisStyle(plotAreaNode, axisType, warpObj) {
    const axisNode = PPTXXmlUtils.getTextByPathList(plotAreaNode, [axisType]);
    if (!axisNode)
        return {};
    const style = {};
    const txPr = PPTXXmlUtils.getTextByPathList(axisNode, ["c:txPr"]);
    if (txPr) {
        const p = PPTXXmlUtils.getTextByPathList(txPr, ["a:p"]);
        if (p) {
            const pPr = PPTXXmlUtils.getTextByPathList(p, ["a:pPr"]);
            if (pPr) {
                const defRPr = PPTXXmlUtils.getTextByPathList(pPr, ["a:defRPr"]);
                if (defRPr) {
                    if (defRPr["attrs"] && defRPr["attrs"]["sz"]) {
                        style.fontSize = parseFloat(defRPr["attrs"]["sz"]) / 100;
                    }
                    const solidFill = PPTXXmlUtils.getTextByPathList(defRPr, ["a:solidFill"]);
                    if (solidFill) {
                        let color = getSolidFill(solidFill, undefined, undefined, warpObj);
                        if (color && !color.startsWith('#')) {
                            color = `#${color}`;
                        }
                        style.color = color;
                    }
                }
            }
        }
    }
    const spPr = PPTXXmlUtils.getTextByPathList(axisNode, ["c:spPr"]);
    if (spPr) {
        const ln = PPTXXmlUtils.getTextByPathList(spPr, ["a:ln"]);
        if (ln) {
            const solidFill = PPTXXmlUtils.getTextByPathList(ln, ["a:solidFill"]);
            if (solidFill) {
                let lineColor = getSolidFill(solidFill, undefined, undefined, warpObj);
                if (lineColor && !lineColor.startsWith('#')) {
                    lineColor = `#${lineColor}`;
                }
                style.lineColor = lineColor;
            }
            if (ln["attrs"] && ln["attrs"]["w"]) {
                style.lineWidth = parseFloat(ln["attrs"]["w"]) / 9525;
            }
        }
    }
    if (axisType === "c:valAx") {
        const majorGridlines = PPTXXmlUtils.getTextByPathList(axisNode, ["c:majorGridlines"]);
        if (majorGridlines) {
            const spPr = PPTXXmlUtils.getTextByPathList(majorGridlines, ["c:spPr"]);
            if (spPr) {
                const ln = PPTXXmlUtils.getTextByPathList(spPr, ["a:ln"]);
                if (ln) {
                    const solidFill = PPTXXmlUtils.getTextByPathList(ln, ["a:solidFill"]);
                    if (solidFill) {
                        let gridlineColor = getSolidFill(solidFill, undefined, undefined, warpObj);
                        if (gridlineColor && !gridlineColor.startsWith('#')) {
                            gridlineColor = `#${gridlineColor}`;
                        }
                        style.gridlineColor = gridlineColor;
                    }
                    if (ln["attrs"] && ln["attrs"]["w"]) {
                        style.gridlineWidth = parseFloat(ln["attrs"]["w"]) / 9525;
                    }
                }
            }
        }
    }
    return style;
}
function getColor(node, clrMap, phClr, warpObj) {
    if (node["a:solidFill"]) {
        return getSolidFill(node["a:solidFill"], clrMap, phClr, warpObj);
    }
    return "";
}
const PPTXStyleUtils = {
    getFillType,
    getShapeFill,
    getFontType,
    getFontColorPr,
    getFontSize,
    getFontBold,
    getFontItalic,
    getFontDecoration,
    getTextHorizontalAlign,
    getTextVerticalAlign,
    getTableBorders,
    getBorder,
    getSlideBackgroundFill,
    getBgGradientFill,
    getBgPicFill,
    getGradientFill,
    getPicFill,
    getPatternFill,
    getLinerGrandient,
    getSolidFill,
    toHex,
    hslToRgb,
    hueToRgb,
    getColorName2Hex,
    getSchemeColorFromTheme,
    extractChartData,
    extractChartTitleStyle,
    extractChartAreaStyle,
    extractChartLegendStyle,
    extractChartAxisStyle,
    setTextByPathList,
    eachElement,
    applyShade,
    applyTint,
    applyLumOff,
    applyLumMod,
    applyHueMod,
    applySatMod,
    rgba2hex,
    getSvgGradient,
    getMiddleStops,
    SVGangle,
    getSvgImagePattern,
    getBase64ImageDimensions,
    getVerticalAlign,
    getContentDir,
    getHorizontalAlign,
    getVerticalMargins,
    getLayoutAndMasterNode,
    getPregraphDir,
    getPregraphMargn,
    getColor
};

const tinycolor = (color, opts) => new tinycolor$2(color, opts);
let is_first_br = false;
function getTextWidth(html) {
    let div = document.createElement('div');
    div.style.position = 'absolute';
    div.style.float = 'left';
    div.style.whiteSpace = 'nowrap';
    div.style.visibility = 'hidden';
    div.innerHTML = html;
    document.body.appendChild(div);
    let width = div.offsetWidth;
    document.body.removeChild(div);
    return width;
}
async function genTextBody(textBodyNode, spNode, slideLayoutSpNode, slideMasterSpNode, type, idx, warpObj, tbl_col_width) {
    let text = "";
    warpObj["slideMasterTextStyles"];
    if (textBodyNode === undefined) {
        return text;
    }
    let anchor = PPTXXmlUtils.getTextByPathList(textBodyNode, ["a:bodyPr", "attrs", "anchor"]);
    if (anchor === undefined) {
        anchor = PPTXXmlUtils.getTextByPathList(slideLayoutSpNode, ["p:txBody", "a:bodyPr", "attrs", "anchor"]);
        if (anchor === undefined) {
            anchor = PPTXXmlUtils.getTextByPathList(slideMasterSpNode, ["p:txBody", "a:bodyPr", "attrs", "anchor"]);
            if (anchor === undefined) {
                anchor = "t";
            }
        }
    }
    if (type === "shape") {
        let shapeType = PPTXXmlUtils.getTextByPathList(spNode, ["p:spPr", "a:prstGeom", "attrs", "prst"]);
        const circularShapes = [
            "ellipse", "ovalCallout", "wedgeEllipseCallout",
            "pie", "pieWedge", "chord", "sector", "arc", "blockArc"
        ];
        if (circularShapes.includes(shapeType) && anchor === "t") {
            anchor = "ctr";
        }
    }
    let bodyPrPadding = getBodyPrPadding(textBodyNode, type, anchor);
    text += bodyPrPadding;
    let pFontStyle = PPTXXmlUtils.getTextByPathList(spNode, ["p:style", "a:fontRef"]);
    let wrapAttr = PPTXXmlUtils.getTextByPathList(textBodyNode["a:bodyPr"], ["attrs", "wrap"]);
    let spAutoFitNode = PPTXXmlUtils.getTextByPathList(textBodyNode["a:bodyPr"], ["a:spAutoFit"]);
    PPTXXmlUtils.getTextByPathList(textBodyNode["a:bodyPr"], ["attrs", "rtlCol"]);
    let isNoWrap = (wrapAttr === "none");
    let isAutoFit = (spAutoFitNode !== undefined);
    let apNode = textBodyNode["a:p"];
    if (apNode.constructor !== Array) {
        apNode = [apNode];
    }
    for (const i of apNode.keys()) {
        let pNode = apNode[i];
        let rNode = pNode["a:r"];
        let fldNode = pNode["a:fld"];
        let brNode = pNode["a:br"];
        if (rNode !== undefined) {
            rNode = (rNode.constructor === Array) ? rNode : [rNode];
        }
        if (rNode !== undefined && fldNode !== undefined) {
            fldNode = (fldNode.constructor === Array) ? fldNode : [fldNode];
            rNode = rNode.concat(fldNode);
        }
        if (rNode !== undefined && brNode !== undefined) {
            is_first_br = true;
            brNode = (brNode.constructor === Array) ? brNode : [brNode];
            brNode.forEach((item, indx) => {
                item.type = "br";
            });
            if (brNode.length > 1) {
                brNode.shift();
            }
            rNode = rNode.concat(brNode);
            rNode.sort((a, b) => {
                return a.attrs.order - b.attrs.order;
            });
        }
        let styleText = "";
        let marginsVer = PPTXStyleUtils.getVerticalMargins(pNode, textBodyNode, type, idx, warpObj, apNode.length, i, anchor);
        if (marginsVer != "") {
            styleText = marginsVer;
        }
        let cssName = "";
        if (styleText in warpObj.styleTable) {
            cssName = warpObj.styleTable[styleText]["name"];
        }
        else {
            cssName = `_css_${(Object.keys(warpObj.styleTable).length + 1)}`;
            warpObj.styleTable[styleText] = {
                "name": cssName,
                "text": styleText
            };
        }
        let prg_width_node = PPTXXmlUtils.getTextByPathList(spNode, ["p:spPr", "a:xfrm", "a:ext", "attrs", "cx"]);
        if (prg_width_node === undefined || prg_width_node === null) {
            prg_width_node = PPTXXmlUtils.getTextByPathList(slideLayoutSpNode, ["p:spPr", "a:xfrm", "a:ext", "attrs", "cx"]);
        }
        if (prg_width_node === undefined || prg_width_node === null) {
            prg_width_node = PPTXXmlUtils.getTextByPathList(slideMasterSpNode, ["p:spPr", "a:xfrm", "a:ext", "attrs", "cx"]);
        }
        let lIns = PPTXXmlUtils.getTextByPathList(textBodyNode, ["a:bodyPr", "attrs", "lIns"]);
        let rIns = PPTXXmlUtils.getTextByPathList(textBodyNode, ["a:bodyPr", "attrs", "rIns"]);
        let lInsPx, rInsPx;
        if (type === "table") {
            lInsPx = lIns ? (parseInt(lIns) * SLIDE_FACTOR$1) : 0;
            rInsPx = rIns ? (parseInt(rIns) * SLIDE_FACTOR$1) : 0;
        }
        else {
            lInsPx = lIns ? (parseInt(lIns) * SLIDE_FACTOR$1) : (type === "diagram" ? 0.04 * 96 : 0.1 * 96);
            rInsPx = rIns ? (parseInt(rIns) * SLIDE_FACTOR$1) : (type === "diagram" ? 0.04 * 96 : 0.1 * 96);
        }
        if (lIns === "0")
            lInsPx = 0;
        if (rIns === "0")
            rInsPx = 0;
        let shapeType = PPTXXmlUtils.getTextByPathList(spNode, ["p:spPr", "a:prstGeom", "attrs", "prst"]);
        const circularShapes = [
            "ellipse", "ovalCallout", "wedgeEllipseCallout",
            "pie", "pieWedge", "chord", "sector", "arc", "blockArc"
        ];
        const isCircularShape = circularShapes.includes(shapeType);
        let sld_prg_width_val = null;
        if (prg_width_node !== undefined && prg_width_node !== null) {
            let parsedWidth = parseInt(prg_width_node);
            if (!isNaN(parsedWidth) && parsedWidth > 0) {
                sld_prg_width_val = Math.round(parsedWidth * SLIDE_FACTOR$1 * 100) / 100;
            }
        }
        if (sld_prg_width_val !== null && warpObj.currentGroupScale) {
            const { scaleX, scaleY } = warpObj.currentGroupScale;
            sld_prg_width_val = Math.round(sld_prg_width_val * scaleX * 100) / 100;
        }
        let sld_prg_width = "";
        if (sld_prg_width_val !== null && !isNoWrap) {
            let availableWidth = sld_prg_width_val - lInsPx - rInsPx;
            if (isCircularShape) {
                availableWidth = availableWidth * 0.95;
            }
            sld_prg_width = `width:${Math.max(0, Math.round(availableWidth * 100) / 100)}px;`;
        }
        else if (sld_prg_width_val === null) {
            sld_prg_width = "width:inherit;";
        }
        let sld_prg_height = "";
        let prg_dir = PPTXStyleUtils.getPregraphDir(pNode, textBodyNode, idx, type, warpObj);
        let isRTL = (prg_dir == "pregraph-rtl");
        let directionStyle = isRTL ? "direction: rtl;" : "direction: ltr;";
        let horizontalAlign = PPTXStyleUtils.getHorizontalAlign(pNode, textBodyNode, idx, type, prg_dir, warpObj, spNode);
        let outerFlexStyle = "";
        if (type !== "table") {
            if (horizontalAlign === "h-right" || horizontalAlign === "h-right-rtl") {
                outerFlexStyle = isRTL ? "justify-content: flex-start;" : "justify-content: flex-end;";
            }
            else if (horizontalAlign === "h-mid") {
                outerFlexStyle = "justify-content: center;";
            }
            else if (horizontalAlign === "h-left-rtl") {
                outerFlexStyle = "justify-content: flex-end;";
            }
            else {
                outerFlexStyle = "justify-content: flex-start;";
            }
        }
        text += `<div style='display: flex;${sld_prg_width}${sld_prg_height}${outerFlexStyle}${directionStyle}' class='slide-prgrph ${horizontalAlign}` + ` ${prg_dir} ` + cssName + "' >";
        let buText_ary = await genBuChar(pNode, i, spNode, textBodyNode, pFontStyle, idx, type, warpObj);
        let isBullate = (buText_ary[0] !== undefined && buText_ary[0] !== null && buText_ary[0] != "") ? true : false;
        let bu_width = (buText_ary[1] !== undefined && buText_ary[1] !== null && isBullate) ? (Number(buText_ary[1]) + Number(buText_ary[2])) : 0;
        if (isRTL && isBullate) ;
        else {
            text += (buText_ary[0] !== undefined) ? buText_ary[0] : "";
        }
        let fontSize = undefined;
        if (rNode !== undefined && rNode.length > 0) {
            fontSize = PPTXStyleUtils.getFontSize(rNode[0], textBodyNode, pFontStyle, 1, type, warpObj);
            if (fontSize && fontSize.endsWith('px')) {
                fontSize = parseFloat(fontSize);
            }
        }
        let margin_ary = PPTXStyleUtils.getPregraphMargn(pNode, idx, type, isBullate, warpObj, fontSize);
        let margin = margin_ary[0];
        let mrgin_val = margin_ary[1];
        if (prg_width_node === undefined && tbl_col_width !== undefined && prg_width_node != 0) {
            prg_width_node = tbl_col_width;
        }
        let prgrph_text = "";
        let total_text_len = 0;
        if (rNode === undefined && pNode !== undefined) {
            let prgr_text = await genSpanElement(pNode, undefined, spNode, textBodyNode, pFontStyle, slideLayoutSpNode, idx, type, 1, warpObj);
            if (isBullate) {
                total_text_len += getTextWidth(prgr_text);
            }
            prgrph_text += prgr_text;
        }
        else if (rNode !== undefined) {
            let previousStyle = {};
            for (const j of rNode.keys()) {
                if (rNode[j]["a:rPr"] && !rNode[j]["a:rPr"]["attrs"] && previousStyle["sz"]) {
                    rNode[j]["a:rPr"]["attrs"] = { "sz": previousStyle["sz"] };
                }
                else if (rNode[j]["a:rPr"] && rNode[j]["a:rPr"]["attrs"] && !rNode[j]["a:rPr"]["attrs"]["sz"] && previousStyle["sz"]) {
                    rNode[j]["a:rPr"]["attrs"]["sz"] = previousStyle["sz"];
                }
                let prgr_text = await genSpanElement(rNode[j], j, spNode, textBodyNode, pFontStyle, slideLayoutSpNode, idx, type, rNode.length, warpObj);
                if (isBullate) {
                    total_text_len += getTextWidth(prgr_text);
                }
                prgrph_text += prgr_text;
                if (rNode[j]["a:rPr"] && rNode[j]["a:rPr"]["attrs"] && rNode[j]["a:rPr"]["attrs"]["sz"]) {
                    previousStyle["sz"] = rNode[j]["a:rPr"]["attrs"]["sz"];
                }
            }
        }
        prg_width_node = parseInt(prg_width_node) * SLIDE_FACTOR$1 - bu_width - mrgin_val;
        prg_width_node = Math.round(prg_width_node * 100) / 100;
        let textContainerWidth = "";
        if (!isAutoFit && !isNoWrap && sld_prg_width_val !== null && !isNaN(sld_prg_width_val) && type !== "table") {
            let availableWidthForTextContainer = sld_prg_width_val - lInsPx - rInsPx;
            if (isCircularShape) {
                availableWidthForTextContainer = availableWidthForTextContainer * 0.95;
            }
            textContainerWidth = `width:${Math.max(0, Math.round(availableWidthForTextContainer * 100) / 100)}px;`;
        }
        if (isRTL && isBullate) {
            textContainerWidth = "";
        }
        let whiteSpaceStyle;
        if (isCircularShape) {
            whiteSpaceStyle = isNoWrap ? "white-space: nowrap;" : "white-space: normal; overflow-wrap: break-word;";
        }
        else {
            whiteSpaceStyle = isNoWrap ? "white-space: nowrap;" : "white-space: pre-wrap;";
        }
        let textAlignStyle = "";
        if (horizontalAlign === "h-mid") {
            textAlignStyle = "text-align: center;";
        }
        else if (horizontalAlign === "h-right" || horizontalAlign === "h-right-rtl") {
            textAlignStyle = "text-align: right;";
        }
        else if (horizontalAlign === "h-left-rtl") {
            textAlignStyle = "text-align: left;";
        }
        else {
            textAlignStyle = "text-align: left;";
        }
        textAlignStyle += " word-break: break-all;";
        let flexStyle = "";
        if (type !== "table") {
            if (horizontalAlign === "h-right" || horizontalAlign === "h-right-rtl") {
                flexStyle = isRTL ? "justify-content: flex-start;" : "justify-content: flex-end;";
            }
            else if (horizontalAlign === "h-mid") {
                flexStyle = "justify-content: center;";
            }
            else if (horizontalAlign === "h-left-rtl") {
                flexStyle = "justify-content: flex-end;";
            }
            else {
                flexStyle = "justify-content: flex-start;";
            }
        }
        text += `<div style='display: flex;${flexStyle}${textContainerWidth}${directionStyle}'>`;
        if (isRTL && isBullate && buText_ary[0] !== undefined) {
            text += buText_ary[0];
        }
        text += `<div style='${styleText}${directionStyle}${whiteSpaceStyle}${margin}${textAlignStyle}'>`;
        text += prgrph_text;
        text += "</div>";
        text += "</div>";
        text += "</div>";
    }
    if (type !== "table") {
        text += "</div>";
    }
    return text;
}
function getBodyPrPadding(textBodyNode, type, anchor) {
    let paddingStyle = "";
    let lIns = PPTXXmlUtils.getTextByPathList(textBodyNode, ["a:bodyPr", "attrs", "lIns"]);
    let tIns = PPTXXmlUtils.getTextByPathList(textBodyNode, ["a:bodyPr", "attrs", "tIns"]);
    let rIns = PPTXXmlUtils.getTextByPathList(textBodyNode, ["a:bodyPr", "attrs", "rIns"]);
    let bIns = PPTXXmlUtils.getTextByPathList(textBodyNode, ["a:bodyPr", "attrs", "bIns"]);
    if (type !== "table") {
        let defaultLIns, defaultTIns, defaultRIns, defaultBIns;
        if (type === "diagram") {
            defaultLIns = 0.04 * 96;
            defaultTIns = 0.02 * 96;
            defaultRIns = 0.04 * 96;
            defaultBIns = 0.02 * 96;
        }
        else {
            defaultLIns = 0.1 * 96;
            defaultTIns = 0.05 * 96;
            defaultRIns = 0.1 * 96;
            defaultBIns = 0.05 * 96;
        }
        let lInsPx = lIns ? (parseInt(lIns) * SLIDE_FACTOR$1).toFixed(2) : defaultLIns.toFixed(2);
        let tInsPx = tIns ? (parseInt(tIns) * SLIDE_FACTOR$1).toFixed(2) : defaultTIns.toFixed(2);
        let rInsPx = rIns ? (parseInt(rIns) * SLIDE_FACTOR$1).toFixed(2) : defaultRIns.toFixed(2);
        let bInsPx = bIns ? (parseInt(bIns) * SLIDE_FACTOR$1).toFixed(2) : defaultBIns.toFixed(2);
        if (lIns === "0")
            lInsPx = "0";
        if (rIns === "0")
            rInsPx = "0";
        let heightStyle = "";
        if (anchor !== "ctr") {
            heightStyle = "height: 100%;";
        }
        paddingStyle = `<div style="padding: ${tInsPx}px ${rInsPx}px ${bInsPx}px ${lInsPx}px; box-sizing: border-box; ${heightStyle}">`;
    }
    return paddingStyle;
}
async function genBuChar(node, i, spNode, textBodyNode, pFontStyle, idx, type, warpObj) {
    warpObj["slideMasterTextStyles"];
    let lstStyle = textBodyNode["a:lstStyle"];
    let rNode = PPTXXmlUtils.getTextByPathList(node, ["a:r"]);
    if (rNode !== undefined && rNode.constructor === Array) {
        rNode = rNode[0];
    }
    let lvl = parseInt(PPTXXmlUtils.getTextByPathList(node["a:pPr"], ["attrs", "lvl"])) + 1;
    if (isNaN(lvl)) {
        lvl = 1;
    }
    let lvlStr = `a:lvl${lvl}pPr`;
    let dfltBultColor, dfltBultSize, bultColor, bultSize, color_tye;
    if (rNode !== undefined) {
        dfltBultColor = await PPTXStyleUtils.getFontColorPr(rNode, spNode, lstStyle, pFontStyle, lvl, idx, type, warpObj);
        color_tye = dfltBultColor[2];
        dfltBultSize = PPTXStyleUtils.getFontSize(rNode, textBodyNode, pFontStyle, lvl, type, warpObj);
    }
    else {
        return "";
    }
    let bullet = "", marRStr = "", marLStr = "", margin_val = 0, font_val = 0;
    let pPrNode = node["a:pPr"];
    let BullNONE = PPTXXmlUtils.getTextByPathList(pPrNode, ["a:buNone"]);
    if (BullNONE !== undefined) {
        return "";
    }
    let buType = "TYPE_NONE";
    let layoutMasterNode = PPTXStyleUtils.getLayoutAndMasterNode(node, idx, type, warpObj);
    let { nodeLaout: pPrNodeLaout, nodeMaster: pPrNodeMaster } = layoutMasterNode;
    let buChar = PPTXXmlUtils.getTextByPathList(pPrNode, ["a:buChar", "attrs", "char"]);
    let buNum = PPTXXmlUtils.getTextByPathList(pPrNode, ["a:buAutoNum", "attrs", "type"]);
    let buPic = PPTXXmlUtils.getTextByPathList(pPrNode, ["a:buBlip"]);
    if (buChar !== undefined) {
        buType = "TYPE_BULLET";
    }
    if (buNum !== undefined) {
        buType = "TYPE_NUMERIC";
    }
    if (buPic !== undefined) {
        buType = "TYPE_BULPIC";
    }
    let buFontSize = PPTXXmlUtils.getTextByPathList(pPrNode, ["a:buSzPts", "attrs", "val"]);
    if (buFontSize === undefined) {
        buFontSize = PPTXXmlUtils.getTextByPathList(pPrNode, ["a:buSzPct", "attrs", "val"]);
        if (buFontSize !== undefined) {
            let prcnt = parseInt(buFontSize) / 100000;
            let dfltBultSizeNoPt = parseInt(dfltBultSize, 10);
            bultSize = `${prcnt * (parseInt(String(dfltBultSizeNoPt)))}px`;
        }
    }
    else {
        bultSize = `${(parseInt(buFontSize) / 100) * FONT_SIZE_FACTOR}px`;
    }
    let buClrNode = PPTXXmlUtils.getTextByPathList(pPrNode, ["a:buClr"]);
    if (buChar === undefined && buNum === undefined && buPic === undefined) {
        if (lstStyle !== undefined) {
            BullNONE = PPTXXmlUtils.getTextByPathList(lstStyle, [lvlStr, "a:buNone"]);
            if (BullNONE !== undefined) {
                return "";
            }
            buType = "TYPE_NONE";
            buChar = PPTXXmlUtils.getTextByPathList(lstStyle, [lvlStr, "a:buChar", "attrs", "char"]);
            buNum = PPTXXmlUtils.getTextByPathList(lstStyle, [lvlStr, "a:buAutoNum", "attrs", "type"]);
            buPic = PPTXXmlUtils.getTextByPathList(lstStyle, [lvlStr, "a:buBlip"]);
            if (buChar !== undefined) {
                buType = "TYPE_BULLET";
            }
            if (buNum !== undefined) {
                buType = "TYPE_NUMERIC";
            }
            if (buPic !== undefined) {
                buType = "TYPE_BULPIC";
            }
            if (buChar !== undefined || buNum !== undefined || buPic !== undefined) {
                pPrNode = lstStyle[lvlStr];
            }
        }
    }
    if (buChar === undefined && buNum === undefined && buPic === undefined) {
        if (pPrNodeLaout !== undefined) {
            BullNONE = PPTXXmlUtils.getTextByPathList(pPrNodeLaout, ["a:buNone"]);
            if (BullNONE !== undefined) {
                return "";
            }
            buType = "TYPE_NONE";
            buChar = PPTXXmlUtils.getTextByPathList(pPrNodeLaout, ["a:buChar", "attrs", "char"]);
            buNum = PPTXXmlUtils.getTextByPathList(pPrNodeLaout, ["a:buAutoNum", "attrs", "type"]);
            buPic = PPTXXmlUtils.getTextByPathList(pPrNodeLaout, ["a:buBlip"]);
            if (buChar !== undefined) {
                buType = "TYPE_BULLET";
            }
            if (buNum !== undefined) {
                buType = "TYPE_NUMERIC";
            }
            if (buPic !== undefined) {
                buType = "TYPE_BULPIC";
            }
        }
        if (buChar === undefined && buNum === undefined && buPic === undefined) {
            if (pPrNodeMaster !== undefined) {
                BullNONE = PPTXXmlUtils.getTextByPathList(pPrNodeMaster, ["a:buNone"]);
                if (BullNONE !== undefined) {
                    return "";
                }
                buType = "TYPE_NONE";
                buChar = PPTXXmlUtils.getTextByPathList(pPrNodeMaster, ["a:buChar", "attrs", "char"]);
                buNum = PPTXXmlUtils.getTextByPathList(pPrNodeMaster, ["a:buAutoNum", "attrs", "type"]);
                buPic = PPTXXmlUtils.getTextByPathList(pPrNodeMaster, ["a:buBlip"]);
                if (buChar !== undefined) {
                    buType = "TYPE_BULLET";
                }
                if (buNum !== undefined) {
                    buType = "TYPE_NUMERIC";
                }
                if (buPic !== undefined) {
                    buType = "TYPE_BULPIC";
                }
            }
        }
    }
    let getRtlVal = PPTXXmlUtils.getTextByPathList(pPrNode, ["attrs", "rtl"]);
    if (getRtlVal === undefined) {
        getRtlVal = PPTXXmlUtils.getTextByPathList(pPrNodeLaout, ["attrs", "rtl"]);
        if (getRtlVal === undefined && type != "shape") {
            getRtlVal = PPTXXmlUtils.getTextByPathList(pPrNodeMaster, ["attrs", "rtl"]);
        }
    }
    let isRTL = false;
    if (getRtlVal !== undefined && getRtlVal == "1") {
        isRTL = true;
    }
    let alignNode = PPTXXmlUtils.getTextByPathList(pPrNode, ["attrs", "algn"]);
    if (alignNode === undefined) {
        alignNode = PPTXXmlUtils.getTextByPathList(pPrNodeLaout, ["attrs", "algn"]);
        if (alignNode === undefined) {
            alignNode = PPTXXmlUtils.getTextByPathList(pPrNodeMaster, ["attrs", "algn"]);
        }
    }
    let indentNode = PPTXXmlUtils.getTextByPathList(pPrNode, ["attrs", "indent"]);
    if (indentNode === undefined) {
        indentNode = PPTXXmlUtils.getTextByPathList(pPrNodeLaout, ["attrs", "indent"]);
        if (indentNode === undefined) {
            indentNode = PPTXXmlUtils.getTextByPathList(pPrNodeMaster, ["attrs", "indent"]);
        }
    }
    let indent = 0;
    if (indentNode !== undefined) {
        indent = parseInt(indentNode) * SLIDE_FACTOR$1;
    }
    let marLNode = PPTXXmlUtils.getTextByPathList(pPrNode, ["attrs", "marL"]);
    if (marLNode === undefined) {
        marLNode = PPTXXmlUtils.getTextByPathList(pPrNodeLaout, ["attrs", "marL"]);
        if (marLNode === undefined) {
            marLNode = PPTXXmlUtils.getTextByPathList(pPrNodeMaster, ["attrs", "marL"]);
        }
    }
    if (marLNode !== undefined) {
        let marginLeft = parseInt(marLNode) * SLIDE_FACTOR$1;
        if (isRTL) {
            marLStr = "padding-right:";
        }
        else {
            marLStr = "padding-left:";
        }
        margin_val = ((marginLeft + indent < 0) ? 0 : (marginLeft + indent));
        marLStr += `${margin_val}px;`;
    }
    let marRNode = PPTXXmlUtils.getTextByPathList(pPrNode, ["attrs", "marR"]);
    if (marRNode === undefined && marLNode === undefined) {
        marRNode = PPTXXmlUtils.getTextByPathList(pPrNodeLaout, ["attrs", "marR"]);
        if (marRNode === undefined) {
            marRNode = PPTXXmlUtils.getTextByPathList(pPrNodeMaster, ["attrs", "marR"]);
        }
    }
    if (marRNode !== undefined) {
        let marginRight = parseInt(marRNode) * SLIDE_FACTOR$1;
        if (isRTL) {
            marLStr = "padding-right:";
        }
        else {
            marLStr = "padding-left:";
        }
        marRStr += `${((marginRight + indent < 0) ? 0 : (marginRight + indent))}px;`;
    }
    if (buClrNode === undefined) {
        buClrNode = PPTXXmlUtils.getTextByPathList(lstStyle, [lvlStr, "a:buClr"]);
    }
    if (buClrNode === undefined) {
        buClrNode = PPTXXmlUtils.getTextByPathList(pPrNodeLaout, ["a:buClr"]);
        if (buClrNode === undefined) {
            buClrNode = PPTXXmlUtils.getTextByPathList(pPrNodeMaster, ["a:buClr"]);
        }
    }
    let defBultColor;
    if (buClrNode !== undefined) {
        defBultColor = PPTXStyleUtils.getSolidFill(buClrNode, undefined, undefined, warpObj);
    }
    if (defBultColor === undefined || defBultColor == "NONE") {
        bultColor = dfltBultColor;
    }
    else {
        bultColor = [defBultColor, "", "solid"];
        color_tye = "solid";
    }
    if (buFontSize === undefined) {
        buFontSize = PPTXXmlUtils.getTextByPathList(pPrNodeLaout, ["a:buSzPts", "attrs", "val"]);
        if (buFontSize === undefined) {
            buFontSize = PPTXXmlUtils.getTextByPathList(pPrNodeLaout, ["a:buSzPct", "attrs", "val"]);
            if (buFontSize !== undefined) {
                let prcnt = parseInt(buFontSize) / 100000;
                let dfltBultSizeNoPt = parseInt(dfltBultSize, 10);
                bultSize = `${prcnt * (parseInt(String(dfltBultSizeNoPt)))}px`;
            }
        }
        else {
            bultSize = `${(parseInt(buFontSize) / 100) * FONT_SIZE_FACTOR}px`;
        }
    }
    if (buFontSize === undefined) {
        buFontSize = PPTXXmlUtils.getTextByPathList(pPrNodeMaster, ["a:buSzPts", "attrs", "val"]);
        if (buFontSize === undefined) {
            buFontSize = PPTXXmlUtils.getTextByPathList(pPrNodeMaster, ["a:buSzPct", "attrs", "val"]);
            if (buFontSize !== undefined) {
                let prcnt = parseInt(buFontSize) / 100000;
                let dfltBultSizeNoPt = parseInt(dfltBultSize, 10);
                bultSize = `${prcnt * (parseInt(String(dfltBultSizeNoPt)))}px`;
            }
        }
        else {
            bultSize = `${(parseInt(buFontSize) / 100) * FONT_SIZE_FACTOR}px`;
        }
    }
    if (buFontSize === undefined) {
        bultSize = dfltBultSize;
    }
    font_val = parseInt(bultSize, 10);
    if (buType == "TYPE_BULLET") {
        let typefaceNode = PPTXXmlUtils.getTextByPathList(pPrNode, ["a:buFont", "attrs", "typeface"]);
        let typeface = "";
        let isWingdingsFont = false;
        if (typefaceNode !== undefined) {
            isWingdingsFont = (typefaceNode == "Wingdings" || typefaceNode == "Wingdings 2" || typefaceNode == "Wingdings 3" || typefaceNode == "Webdings");
            typeface = `font-family: ${typefaceNode}`;
        }
        bullet = `<div style='${typeface};` +
            marLStr + marRStr +
            `font-size:${bultSize};`;
        if (color_tye == "solid") {
            if (bultColor[0] !== undefined && bultColor[0] != "") {
                let bulletColorValue = bultColor[0];
                if (bulletColorValue.length === 8) {
                    let colorObj = tinycolor(bulletColorValue);
                    bulletColorValue = colorObj.toRgbString();
                }
                else {
                    bulletColorValue = `#${bulletColorValue}`;
                }
                bullet += `color:${bulletColorValue}; `;
            }
            if (bultColor[1] !== undefined && bultColor[1] != "" && bultColor[1] != ";") {
                bullet += `text-shadow:${bultColor[1]};`;
            }
        }
        else if (color_tye == "pattern" || color_tye == "pic" || color_tye == "gradient") {
            if (color_tye == "pattern") {
                bullet += `background:${bultColor[0][0]};`;
                if (bultColor[0][1] !== null && bultColor[0][1] !== undefined && bultColor[0][1] != "") {
                    bullet += `background-size:${bultColor[0][1]};`;
                }
                if (bultColor[0][2] !== null && bultColor[0][2] !== undefined && bultColor[0][2] != "") {
                    bullet += `background-position:${bultColor[0][2]};`;
                }
            }
            else if (color_tye == "pic") {
                bullet += `${bultColor[0]};`;
            }
            else if (color_tye == "gradient") {
                let colorAry = bultColor[0].color;
                if (!Array.isArray(colorAry)) {
                    colorAry = colorAry ? [colorAry] : [];
                }
                let rot = bultColor[0].rot;
                bullet += `background: linear-gradient(${rot}deg,`;
                for (const i of colorAry.keys()) {
                    if (i == colorAry.length - 1) {
                        bullet += `#${colorAry[i]});`;
                    }
                    else {
                        bullet += `#${colorAry[i]}, `;
                    }
                }
            }
            bullet += `-webkit-background-clip: text;background-clip: text;color: transparent;`;
            if (bultColor[1].border !== undefined && bultColor[1].border !== "") {
                bullet += `-webkit-text-stroke: ${bultColor[1].border};`;
            }
            if (bultColor[1].effcts !== undefined && bultColor[1].effcts !== "") {
                bullet += `filter: ${bultColor[1].effcts};`;
            }
        }
        if (isRTL) {
            bullet += "white-space: nowrap ;direction:rtl";
        }
        let isIE11 = !!window.MSInputMethodContext && !!document.documentMode;
        let htmlBu = buChar;
        let useUnicodeFont = false;
        if (!isIE11 && !isWingdingsFont) {
            htmlBu = getHtmlBullet(typefaceNode, buChar);
            useUnicodeFont = (htmlBu !== buChar);
        }
        if (useUnicodeFont && isWingdingsFont && typefaceNode !== undefined) {
            bullet = bullet.replace(/font-family:\s*(Wingdings|Wingdings\s*2|Wingdings\s*3|Webdings)\s*/gi, "font-family: Arial, sans-serif");
        }
        bullet += `display: flex; align-items: center;'><div>${htmlBu}</div></div>`;
    }
    else if (buType == "TYPE_NUMERIC") {
        if (!warpObj.bulletCounter) {
            warpObj.bulletCounter = {};
        }
        const bulletKey = `${buNum}_${lvl}`;
        if (!warpObj.bulletCounter[bulletKey]) {
            warpObj.bulletCounter[bulletKey] = {
                index: 0,
                type: buNum,
                level: lvl
            };
        }
        warpObj.bulletCounter[bulletKey].index++;
        const bulletIndex = warpObj.bulletCounter[bulletKey].index;
        const bulletText = getNumTypeNum(buNum, bulletIndex);
        bullet = `<div style='${marLStr}${marRStr}`;
        if (bultColor && bultColor[0] !== undefined && bultColor[0] != "") {
            let bulletNumColorValue = bultColor[0];
            if (bulletNumColorValue.length === 8) {
                let colorObj = tinycolor(bulletNumColorValue);
                bulletNumColorValue = colorObj.toRgbString();
            }
            else {
                bulletNumColorValue = `#${bulletNumColorValue}`;
            }
            bullet += `color:${bulletNumColorValue};`;
        }
        bullet += `font-size:${bultSize};`;
        if (isRTL) {
            bullet += "white-space: nowrap ;direction:rtl;";
        }
        else {
            bullet += "white-space: nowrap ;direction:ltr;";
        }
        bullet += `display: flex; align-items: center;'><div>${bulletText}</div></div>`;
    }
    else if (buType == "TYPE_BULPIC") {
        let buPicId = PPTXXmlUtils.getTextByPathList(buPic, ["a:blip", "attrs", "r:embed"]);
        let buImg;
        if (buPicId !== undefined) {
            let imgPath = (warpObj["slideResObj"][buPicId] !== undefined) ? warpObj["slideResObj"][buPicId]["target"] : undefined;
            if (imgPath === undefined) {
                buImg = "";
            }
            else {
                let imgFile = warpObj["zip"].file(imgPath);
                if (imgFile === null) {
                    buImg = "";
                }
                else {
                    let imgArrayBuffer = await imgFile.async("arraybuffer");
                    let imgExt = imgPath.split(".").pop();
                    let imgMimeType = PPTXXmlUtils.getMimeType(imgExt);
                    buImg = `<img src='data:${imgMimeType};base64,` + PPTXXmlUtils.base64ArrayBuffer(imgArrayBuffer) + "' style='width: 100%;'/>";
                }
            }
        }
        if (buPicId === undefined) {
            buImg = "&#8227;";
        }
        bullet = `<div style='${marLStr}${marRStr}` +
            `width:${bultSize};display: flex; align-items: center;`;
        if (isRTL) {
            bullet += "white-space: nowrap ;direction:rtl;";
        }
        bullet += `'>${buImg}  </div>`;
    }
    return [bullet, margin_val, font_val];
}
function getHtmlBullet(typefaceNode, buChar) {
    switch (buChar) {
        case "§":
            return "&#9632;";
        case "q":
            return "&#10065;";
        case "v":
            return "&#10070;";
        case "Ø":
            return "&#11162;";
        case "ü":
            return "&#10004;";
        case "o":
            return "&#9679;";
        case "O":
            return "&#9675;";
        case "a":
            return "&#9650;";
        case "A":
            return "&#9651;";
        case "b":
            return "&#9660;";
        case "B":
            return "&#9661;";
        case "c":
            return "&#9654;";
        case "C":
            return "&#9655;";
        case "d":
            return "&#9664;";
        case "D":
            return "&#9665;";
        case "e":
            return "&#9670;";
        case "E":
            return "&#9671;";
        case "f":
            return "&#10003;";
        case "F":
            return "&#10007;";
        case "g":
            return "&#10002;";
        case "G":
            return "&#10008;";
        case "h":
            return "&#9899;";
        case "H":
            return "&#9734;";
        case "i":
            return "&#10052;";
        case "I":
            return "&#10053;";
        case "j":
            return "&#10022;";
        case "J":
            return "&#10023;";
        case "k":
            return "&#10016;";
        case "K":
            return "&#10024;";
        case "l":
            return "&#10038;";
        case "L":
            return "&#10039;";
        case "m":
            return "&#10017;";
        case "M":
            return "&#9993;";
        case "n":
            return "&#10084;";
        case "N":
            return "&#9829;";
        case "p":
            return "&#9830;";
        case "P":
            return "&#9826;";
        case "r":
            return "&#9827;";
        case "R":
            return "&#9827;";
        case "s":
            return "&#9824;";
        case "S":
            return "&#9824;";
        case "t":
            return "&#9828;";
        case "T":
            return "&#9825;";
        case "u":
            return "&#9829;";
        case "U":
            return "&#9825;";
        case "w":
            return "&#10071;";
        case "W":
            return "&#10071;";
        case "x":
            return "&#10062;";
        case "X":
            return "&#10063;";
        case "y":
            return "&#10064;";
        case "Y":
            return "&#10064;";
        case "z":
            return "&#10061;";
        case "Z":
            return "&#10061;";
        default:
            if (typefaceNode == "Wingdings" || typefaceNode == "Wingdings 2" || typefaceNode == "Wingdings 3" || typefaceNode == "Webdings") {
                let wingCharCode = getDingbatToUnicode(typefaceNode, buChar);
                if (wingCharCode !== null) {
                    return `&#${wingCharCode};`;
                }
            }
            return `&#${(buChar.charCodeAt(0))};`;
    }
}
function getDingbatToUnicode(typefaceNode, buChar) {
    if (DINGBAT_UNICODE) {
        let dingbat_code = buChar.codePointAt(0) & 0xFFF;
        let char_unicode = null;
        let len = DINGBAT_UNICODE.length;
        let i = 0;
        while (len--) {
            let item = DINGBAT_UNICODE[i];
            if (item.f == typefaceNode && item.code == dingbat_code) {
                char_unicode = item.unicode;
                break;
            }
            i++;
        }
        return char_unicode;
    }
}
function alphaNumeric(num, upperLower) {
    num = Number(num) - 1;
    let aNum = "";
    if (upperLower == "upperCase") {
        aNum = (((num / 26 >= 1) ? String.fromCharCode(num / 26 + 64) : '') + String.fromCharCode(num % 26 + 65)).toUpperCase();
    }
    else if (upperLower == "lowerCase") {
        aNum = (((num / 26 >= 1) ? String.fromCharCode(num / 26 + 64) : '') + String.fromCharCode(num % 26 + 65)).toLowerCase();
    }
    return aNum;
}
function hebrewAlphaNumeric(num) {
    num = Number(num) - 1;
    const hebrewLetters = [
        'א', 'ב', 'ג', 'ד', 'ה', 'ו', 'ז', 'ח', 'ט',
        'י', 'כ', 'ל', 'מ', 'נ', 'ס', 'ע', 'פ', 'צ',
        'ק', 'ר', 'ש', 'ת'
    ];
    const hebrewLength = hebrewLetters.length;
    if (num < hebrewLength) {
        return hebrewLetters[num];
    }
    else if (num < hebrewLength * (hebrewLength + 1)) {
        const first = Math.floor(num / hebrewLength);
        const second = num % hebrewLength;
        return hebrewLetters[first] + hebrewLetters[second];
    }
    else {
        const third = num % hebrewLength;
        const remaining = Math.floor(num / hebrewLength);
        const second = remaining % hebrewLength;
        const first = Math.floor(remaining / hebrewLength);
        return hebrewLetters[first] + hebrewLetters[second] + hebrewLetters[third];
    }
}
function archaicNumbers(arr) {
    arr.slice().sort((a, b) => { return b[1].length - a[1].length; });
    return {
        format: (n) => {
            let ret = '';
            for (const item of arr) {
                let num = item[0];
                if (parseInt(num) > 0) {
                    for (; n >= num; n -= num)
                        ret += item[1];
                }
                else {
                    ret = ret.replace(num, item[1]);
                }
            }
            return ret;
        }
    };
}
function romanize(num) {
    if (!+num)
        return false;
    let digits = String(+num).split(""), key = ["", "C", "CC", "CCC", "CD", "D", "DC", "DCC", "DCCC", "CM",
        "", "X", "XX", "XXX", "XL", "L", "LX", "LXX", "LXXX", "XC",
        "", "I", "II", "III", "IV", "V", "VI", "VII", "VIII", "IX"], roman = "", i = 3;
    while (i--)
        roman = (key[+digits.pop() + (i * 10)] || "") + roman;
    return Array(+digits.join("") + 1).join("M") + roman;
}
archaicNumbers([
    [1000, ''],
    [400, 'ת'],
    [300, 'ש'],
    [200, 'ר'],
    [100, 'ק'],
    [90, 'צ'],
    [80, 'פ'],
    [70, 'ע'],
    [60, 'ס'],
    [50, 'נ'],
    [40, 'מ'],
    [30, 'ל'],
    [20, 'כ'],
    [10, 'י'],
    [9, 'ט'],
    [8, 'ח'],
    [7, 'ז'],
    [6, 'ו'],
    [5, 'ה'],
    [4, 'ד'],
    [3, 'ג'],
    [2, 'ב'],
    [1, 'א'],
    [/יה/, 'ט״ו'],
    [/יו/, 'ט״ז'],
    [/([א-ת])([א-ת])$/, '$1״$2'],
    [/^([א-ת])$/, "$1׳"]
]);
function getNumTypeNum(numTyp, num) {
    let rtrnNum = "";
    switch (numTyp) {
        case "arabicPeriod":
            rtrnNum = `${num}. `;
            break;
        case "arabicParenR":
            rtrnNum = `${num}) `;
            break;
        case "alphaLcParenR":
            rtrnNum = `${alphaNumeric(num, "lowerCase")}) `;
            break;
        case "alphaLcPeriod":
            rtrnNum = `${alphaNumeric(num, "lowerCase")}. `;
            break;
        case "alphaUcParenR":
            rtrnNum = `${alphaNumeric(num, "upperCase")}) `;
            break;
        case "alphaUcPeriod":
            rtrnNum = `${alphaNumeric(num, "upperCase")}. `;
            break;
        case "romanUcPeriod":
            rtrnNum = `${romanize(num)}. `;
            break;
        case "romanLcParenR":
            rtrnNum = `${romanize(num)}) `;
            break;
        case "hebrew2Minus":
            rtrnNum = `${hebrewAlphaNumeric(num)}-`;
            break;
        default:
            rtrnNum = num;
    }
    return rtrnNum;
}
async function genSpanElement(node, rIndex, pNode, textBodyNode, pFontStyle, slideLayoutSpNode, idx, type, rNodeLength, warpObj, isBullate) {
    let text_style = "";
    let lstStyle = textBodyNode["a:lstStyle"];
    let slideMasterTextStyles = warpObj["slideMasterTextStyles"];
    let text = node["a:t"];
    let rtlColAttr = PPTXXmlUtils.getTextByPathList(textBodyNode, ["a:bodyPr", "attrs", "rtlCol"]);
    let wrapAttr = PPTXXmlUtils.getTextByPathList(textBodyNode, ["a:bodyPr", "attrs", "wrap"]);
    let isRTLCol = (rtlColAttr === "1" && wrapAttr === undefined);
    let openElemnt = "<span";
    let closeElemnt = "</span>";
    let styleText = "";
    if (text === undefined && node["type"] !== undefined) {
        if (is_first_br) {
            is_first_br = false;
            return "<span class='line-break-br' ></span>";
        }
        styleText += "display: block;";
    }
    else {
        is_first_br = true;
    }
    if (typeof text !== 'string') {
        text = PPTXXmlUtils.getTextByPathList(node, ["a:fld", "a:t"]);
        if (typeof text !== 'string') {
            text = "&nbsp;";
        }
    }
    let pPrNode = pNode["a:pPr"];
    let lvl = 1;
    let lvlNode = PPTXXmlUtils.getTextByPathList(pPrNode, ["attrs", "lvl"]);
    if (lvlNode !== undefined) {
        lvl = parseInt(lvlNode) + 1;
    }
    let layoutMasterNode = PPTXStyleUtils.getLayoutAndMasterNode(pNode, idx, type, warpObj);
    let { nodeLaout: pPrNodeLaout, nodeMaster: pPrNodeMaster } = layoutMasterNode;
    let lang = PPTXXmlUtils.getTextByPathList(node, ["a:rPr", "attrs", "lang"]);
    let isRtlLan = (lang !== undefined && RTL_LANGS_ARRAY.indexOf(lang) !== -1) ? true : false;
    let getRtlVal = PPTXXmlUtils.getTextByPathList(pPrNode, ["attrs", "rtl"]);
    if (getRtlVal === undefined) {
        getRtlVal = PPTXXmlUtils.getTextByPathList(pPrNodeLaout, ["attrs", "rtl"]);
        if (getRtlVal === undefined && type != "shape") {
            getRtlVal = PPTXXmlUtils.getTextByPathList(pPrNodeMaster, ["attrs", "rtl"]);
        }
    }
    let linkID = PPTXXmlUtils.getTextByPathList(node, ["a:rPr", "a:hlinkClick", "attrs", "r:id"]);
    let linkTooltip = "";
    let defLinkClr;
    if (linkID !== undefined) {
        const tip = PPTXXmlUtils.getTextByPathList(node, ["a:rPr", "a:hlinkClick", "attrs", "tooltip"]);
        if (tip !== undefined) {
            linkTooltip = `title='${tip}'`;
        }
        defLinkClr = PPTXStyleUtils.getSchemeColorFromTheme("a:hlink", undefined, undefined, warpObj);
    }
    else {
        linkID = PPTXXmlUtils.getTextByPathList(node, ["a:rPr", "a:hlinkHover", "attrs", "r:id"]);
        if (linkID !== undefined) {
            const tip = PPTXXmlUtils.getTextByPathList(node, ["a:rPr", "a:hlinkHover", "attrs", "tooltip"]);
            if (tip !== undefined) {
                linkTooltip = `title='${tip}'`;
            }
            defLinkClr = PPTXStyleUtils.getSchemeColorFromTheme("a:hlink", undefined, undefined, warpObj);
        }
    }
    if (linkID !== undefined) {
        let linkClrNode = PPTXXmlUtils.getTextByPathList(node, ["a:rPr", "a:solidFill"]);
        PPTXStyleUtils.getSolidFill(linkClrNode, undefined, undefined, warpObj);
    }
    let fontClrPr = await PPTXStyleUtils.getFontColorPr(node, pNode, lstStyle, pFontStyle, lvl, idx, type, warpObj);
    let fontClrType = fontClrPr[2];
    if (fontClrType == "solid") {
        if (linkID === undefined && fontClrPr[0] !== undefined && fontClrPr[0] != "") {
            let colorValue = fontClrPr[0];
            if (colorValue.length === 8) {
                let colorObj = tinycolor(colorValue);
                colorValue = colorObj.toRgbString();
            }
            else {
                colorValue = `#${colorValue}`;
            }
            styleText += `color: ${colorValue};`;
        }
        else if (linkID !== undefined && defLinkClr !== undefined) {
            styleText += `color: #${defLinkClr};`;
        }
        if (fontClrPr[1] !== undefined && fontClrPr[1] != "" && fontClrPr[1] != ";") {
            styleText += `text-shadow:${fontClrPr[1]};`;
        }
        if (fontClrPr[3] !== undefined && fontClrPr[3] != "") {
            let highlightColorValue = fontClrPr[3];
            if (highlightColorValue.length === 8) {
                let colorObj = tinycolor(highlightColorValue);
                highlightColorValue = colorObj.toRgbString();
            }
            else {
                highlightColorValue = `#${highlightColorValue}`;
            }
            styleText += `background-color: ${highlightColorValue};`;
        }
    }
    else if (fontClrType == "pattern" || fontClrType == "pic" || fontClrType == "gradient") {
        if (fontClrType == "pattern") {
            styleText += `background:${fontClrPr[0][0]};`;
            if (fontClrPr[0][1] !== null && fontClrPr[0][1] !== undefined && fontClrPr[0][1] != "") {
                styleText += `background-size:${fontClrPr[0][1]};`;
            }
            if (fontClrPr[0][2] !== null && fontClrPr[0][2] !== undefined && fontClrPr[0][2] != "") {
                styleText += `background-position:${fontClrPr[0][2]};`;
            }
        }
        else if (fontClrType == "pic") {
            styleText += `${fontClrPr[0]};`;
        }
        else if (fontClrType == "gradient") {
            let colorAry = fontClrPr[0].color;
            if (!Array.isArray(colorAry)) {
                colorAry = colorAry ? [colorAry] : [];
            }
            let rot = fontClrPr[0].rot;
            styleText += `background: linear-gradient(${rot}deg,`;
            for (const i of colorAry.keys()) {
                if (i == colorAry.length - 1) {
                    styleText += `#${colorAry[i]});`;
                }
                else {
                    styleText += `#${colorAry[i]}, `;
                }
            }
        }
        styleText += `-webkit-background-clip: text;background-clip: text;color: transparent;`;
        if (fontClrPr[1].border !== undefined && fontClrPr[1].border !== "") {
            styleText += `-webkit-text-stroke: ${fontClrPr[1].border};`;
        }
        if (fontClrPr[1].effcts !== undefined && fontClrPr[1].effcts !== "") {
            styleText += `filter: ${fontClrPr[1].effcts};`;
        }
    }
    let font_size = PPTXStyleUtils.getFontSize(node, textBodyNode, pFontStyle, lvl, type, warpObj);
    text_style += `font-size:${font_size};` +
        "font-family:" + PPTXStyleUtils.getFontType(node, type, warpObj, pFontStyle) + ";" +
        "font-weight:" + PPTXStyleUtils.getFontBold(node, type, slideMasterTextStyles) + ";" +
        "font-style:" + PPTXStyleUtils.getFontItalic(node, type, slideMasterTextStyles) + ";" +
        "text-decoration:" + PPTXStyleUtils.getFontDecoration(node, type, slideMasterTextStyles) + ";" +
        "text-align:" + PPTXStyleUtils.getTextHorizontalAlign(node, pNode, type, warpObj) + ";" +
        "vertical-align:" + PPTXStyleUtils.getTextVerticalAlign(node, type, slideMasterTextStyles) + ";";
    text_style += styleText;
    if (isRtlLan) {
        styleText += "direction:rtl;";
    }
    else {
        styleText += "direction:ltr;";
    }
    let highlight = PPTXXmlUtils.getTextByPathList(node, ["a:rPr", "a:highlight"]);
    if (highlight !== undefined) {
        let highlightColor = PPTXStyleUtils.getSolidFill(highlight, undefined, undefined, warpObj);
        if (highlightColor !== undefined && highlightColor != "") {
            if (highlightColor.length === 8) {
                let colorObj = tinycolor(highlightColor);
                highlightColor = colorObj.toRgbString();
            }
            else {
                highlightColor = `#${highlightColor}`;
            }
            styleText += `background-color:${highlightColor};`;
        }
    }
    let spcNode = PPTXXmlUtils.getTextByPathList(node, ["a:rPr", "attrs", "spc"]);
    if (spcNode === undefined) {
        spcNode = PPTXXmlUtils.getTextByPathList(pPrNodeLaout, ["a:defRPr", "attrs", "spc"]);
        if (spcNode === undefined) {
            spcNode = PPTXXmlUtils.getTextByPathList(pPrNodeMaster, ["a:defRPr", "attrs", "spc"]);
        }
    }
    if (spcNode !== undefined) {
        let ltrSpc = parseInt(spcNode) / 100;
        styleText += `letter-spacing: ${ltrSpc}px;`;
    }
    let capNode = PPTXXmlUtils.getTextByPathList(node, ["a:rPr", "attrs", "cap"]);
    if (capNode === undefined) {
        capNode = PPTXXmlUtils.getTextByPathList(pPrNodeLaout, ["a:defRPr", "attrs", "cap"]);
        if (capNode === undefined) {
            capNode = PPTXXmlUtils.getTextByPathList(pPrNodeMaster, ["a:defRPr", "attrs", "cap"]);
        }
    }
    if (capNode == "small" || capNode == "all") {
        styleText += "text-transform: uppercase";
    }
    let cssName = "";
    if (styleText in warpObj.styleTable) {
        cssName = warpObj.styleTable[styleText]["name"];
    }
    else {
        cssName = `_css_${(Object.keys(warpObj.styleTable).length + 1)}`;
        warpObj.styleTable[styleText] = {
            "name": cssName,
            "text": styleText
        };
    }
    let linkColorSyle = "";
    if (fontClrType == "solid" && linkID !== undefined) {
        if (defLinkClr !== undefined) {
            linkColorSyle = `style='color: #${defLinkClr};'`;
        }
    }
    if (linkID !== undefined && linkID != "") {
        const linkRes = warpObj["slideResObj"][linkID];
        let linkURL = linkRes && linkRes.target ? linkRes.target : "";
        const linkType = linkRes && linkRes.type ? linkRes.type : "";
        let linkTargetAttr = " target='_blank'";
        if (linkType === "slide") {
            const m = linkURL.match(/slide(\d+)\.xml$/i);
            if (m) {
                linkURL = `#slide-${m[1]}`;
                linkTargetAttr = "";
            }
        }
        linkURL = PPTXXmlUtils.escapeHtml(linkURL);
        let processedText = text
            .replace(/\t/g, '&nbsp;&nbsp;&nbsp;&nbsp;')
            .replace(/\n/g, "<br>")
            .replace(/  +/g, (spaces) => '&nbsp;'.repeat(spaces.length));
        if (isRTLCol) {
            processedText = processedText.split(/\s+/).filter((word) => word.length > 0).join("<br>");
        }
        return openElemnt + ` class='text-block ${cssName}' style='` + text_style + `'><a href='${linkURL}' ` + linkColorSyle + `  ${linkTooltip}${linkTargetAttr}>` +
            processedText + "</a>" + closeElemnt;
    }
    else {
        let processedText = text
            .replace(/\t/g, '&nbsp;&nbsp;&nbsp;&nbsp;')
            .replace(/\n/g, "<br>")
            .replace(/  +/g, (spaces) => '&nbsp;'.repeat(spaces.length));
        if (isRTLCol) {
            processedText = processedText.split(/\s+/).filter((word) => word.length > 0).join("<br>");
        }
        return openElemnt + ` class='text-block ${cssName}' style='` + text_style + "'>" + processedText + closeElemnt;
    }
}
async function genTable(node, warpObj, shapeType) {
    let order = node["attrs"]["order"];
    let tableNode = PPTXXmlUtils.getTextByPathList(node, ["a:graphic", "a:graphicData", "a:tbl"]);
    let xfrmNode = PPTXXmlUtils.getTextByPathList(node, ["p:xfrm"]);
    let workingXfrmNode = xfrmNode;
    if (shapeType === 'group-abs' && warpObj.currentGroupScale && xfrmNode) {
        const { scaleX, scaleY, childX, childY } = warpObj.currentGroupScale;
        workingXfrmNode = JSON.parse(JSON.stringify(xfrmNode));
        if (xfrmNode['a:ext'] && xfrmNode['a:ext'].attrs) {
            const originalCx = parseInt(xfrmNode['a:ext'].attrs.cx);
            const originalCy = parseInt(xfrmNode['a:ext'].attrs.cy);
            workingXfrmNode['a:ext'].attrs.cx = Math.round(originalCx * scaleX);
            workingXfrmNode['a:ext'].attrs.cy = Math.round(originalCy * scaleY);
        }
        if (xfrmNode['a:off'] && xfrmNode['a:off'].attrs) {
            const originalOffX = parseInt(xfrmNode['a:off'].attrs.x);
            const originalOffY = parseInt(xfrmNode['a:off'].attrs.y);
            const relativeX = originalOffX - (childX / SLIDE_FACTOR$1);
            const relativeY = originalOffY - (childY / SLIDE_FACTOR$1);
            workingXfrmNode['a:off'].attrs.x = Math.round(childX / SLIDE_FACTOR$1 + relativeX * scaleX);
            workingXfrmNode['a:off'].attrs.y = Math.round(childY / SLIDE_FACTOR$1 + relativeY * scaleY);
        }
    }
    let getTblPr = PPTXXmlUtils.getTextByPathList(node, ["a:graphic", "a:graphicData", "a:tbl", "a:tblPr"]);
    let getColsGrid = PPTXXmlUtils.getTextByPathList(node, ["a:graphic", "a:graphicData", "a:tbl", "a:tblGrid", "a:gridCol"]);
    let tblDir = "";
    let firstRowAttr = getTblPr["attrs"]["firstRow"];
    let firstColAttr = getTblPr["attrs"]["firstCol"];
    let lastRowAttr = getTblPr["attrs"]["lastRow"];
    let lastColAttr = getTblPr["attrs"]["lastCol"];
    let bandRowAttr = getTblPr["attrs"]["bandRow"];
    let bandColAttr = getTblPr["attrs"]["bandCol"];
    let tblStylAttrObj = {
        isFrstRowAttr: (firstRowAttr !== undefined && firstRowAttr == "1") ? 1 : 0,
        isFrstColAttr: (firstColAttr !== undefined && firstColAttr == "1") ? 1 : 0,
        isLstRowAttr: (lastRowAttr !== undefined && lastRowAttr == "1") ? 1 : 0,
        isLstColAttr: (lastColAttr !== undefined && lastColAttr == "1") ? 1 : 0,
        isBandRowAttr: (bandRowAttr !== undefined && bandRowAttr == "1") ? 1 : 0,
        isBandColAttr: (bandColAttr !== undefined && bandColAttr == "1") ? 1 : 0
    };
    let thisTblStyle;
    let tbleStyleId = getTblPr["a:tableStyleId"];
    if (tbleStyleId !== undefined) {
        let tbleStylList = warpObj.tableStyles["a:tblStyleLst"]["a:tblStyle"];
        if (tbleStylList !== undefined) {
            if (tbleStylList.constructor === Array) {
                for (const item of tbleStylList) {
                    if (item["attrs"]["styleId"] == tbleStyleId) {
                        thisTblStyle = item;
                    }
                }
            }
            else {
                if (tbleStylList["attrs"]["styleId"] == tbleStyleId) {
                    thisTblStyle = tbleStylList;
                }
            }
        }
    }
    if (thisTblStyle !== undefined) {
        thisTblStyle["tblStylAttrObj"] = tblStylAttrObj;
        warpObj["thisTbiStyle"] = thisTblStyle;
    }
    let tblStyl = PPTXXmlUtils.getTextByPathList(thisTblStyle, ["a:wholeTbl", "a:tcStyle"]);
    let tblBorderStyl = PPTXXmlUtils.getTextByPathList(tblStyl, ["a:tcBdr"]);
    let tbl_borders = "";
    if (tblBorderStyl !== undefined) {
        tbl_borders = PPTXStyleUtils.getTableBorders(tblBorderStyl, warpObj);
    }
    let tbl_bgcolor = "";
    let tbl_bgFillschemeClr = PPTXXmlUtils.getTextByPathList(thisTblStyle, ["a:tblBg", "a:fillRef"]);
    if (tbl_bgFillschemeClr !== undefined) {
        tbl_bgcolor = PPTXStyleUtils.getSolidFill(tbl_bgFillschemeClr, undefined, undefined, warpObj);
    }
    if (tbl_bgFillschemeClr === undefined) {
        tbl_bgFillschemeClr = PPTXXmlUtils.getTextByPathList(thisTblStyle, ["a:wholeTbl", "a:tcStyle", "a:fill", "a:solidFill"]);
        tbl_bgcolor = PPTXStyleUtils.getSolidFill(tbl_bgFillschemeClr, undefined, undefined, warpObj);
    }
    if (tbl_bgcolor !== "" && typeof tbl_bgcolor === 'string') {
        if (tbl_bgcolor.length === 8) {
            let colorObj = tinycolor(tbl_bgcolor);
            tbl_bgcolor = colorObj.toRgbString();
        }
        else {
            tbl_bgcolor = `#${tbl_bgcolor}`;
        }
        tbl_bgcolor = `background-color: ${tbl_bgcolor};`;
    }
    let tableHtml = `<table ${tblDir} style='border-collapse: collapse;` +
        PPTXXmlUtils.getPosition(workingXfrmNode, node, undefined, undefined, shapeType) +
        PPTXXmlUtils.getSize(workingXfrmNode, undefined, undefined) +
        ` z-index: ${order};` +
        tbl_borders + `;${tbl_bgcolor}'>`;
    let trNodes = tableNode["a:tr"];
    if (trNodes.constructor !== Array) {
        trNodes = [trNodes];
    }
    let rowSpanAry = [];
    for (const i of trNodes.keys()) {
        let rowHeightParam = trNodes[i]["attrs"]["h"];
        let rowHeight = 0;
        let rowsStyl = "";
        if (rowHeightParam !== undefined) {
            rowHeight = parseInt(rowHeightParam) * SLIDE_FACTOR$1;
            rowHeight = Math.round(rowHeight * 100) / 100;
            rowsStyl += `height:${rowHeight}px;`;
        }
        let fillColor = "";
        let row_borders = "";
        let fontClrPr = "";
        let fontWeight = "";
        if (thisTblStyle !== undefined && thisTblStyle["a:wholeTbl"] !== undefined) {
            let bgFillschemeClr = PPTXXmlUtils.getTextByPathList(thisTblStyle, ["a:wholeTbl", "a:tcStyle", "a:fill", "a:solidFill"]);
            if (bgFillschemeClr !== undefined) {
                let local_fillColor = PPTXStyleUtils.getSolidFill(bgFillschemeClr, undefined, undefined, warpObj);
                if (local_fillColor !== undefined) {
                    fillColor = local_fillColor;
                }
            }
            let rowTxtStyl = PPTXXmlUtils.getTextByPathList(thisTblStyle, ["a:wholeTbl", "a:tcTxStyle"]);
            if (rowTxtStyl !== undefined) {
                let local_fontColor = PPTXStyleUtils.getSolidFill(rowTxtStyl, undefined, undefined, warpObj);
                if (local_fontColor !== undefined) {
                    fontClrPr = local_fontColor;
                }
                let local_fontWeight = ((PPTXXmlUtils.getTextByPathList(rowTxtStyl, ["attrs", "b"]) == "on") ? "bold" : "");
                if (local_fontWeight != "") {
                    fontWeight = local_fontWeight;
                }
            }
        }
        if (i == 0 && tblStylAttrObj["isFrstRowAttr"] == 1 && thisTblStyle !== undefined) {
            let bgFillschemeClr = PPTXXmlUtils.getTextByPathList(thisTblStyle, ["a:firstRow", "a:tcStyle", "a:fill", "a:solidFill"]);
            if (bgFillschemeClr !== undefined) {
                let local_fillColor = PPTXStyleUtils.getSolidFill(bgFillschemeClr, undefined, undefined, warpObj);
                if (local_fillColor !== undefined) {
                    fillColor = local_fillColor;
                }
            }
            let borderStyl = PPTXXmlUtils.getTextByPathList(thisTblStyle, ["a:firstRow", "a:tcStyle", "a:tcBdr"]);
            if (borderStyl !== undefined) {
                let local_row_borders = PPTXStyleUtils.getTableBorders(borderStyl, warpObj);
                if (local_row_borders != "") {
                    row_borders = local_row_borders;
                }
            }
            let rowTxtStyl = PPTXXmlUtils.getTextByPathList(thisTblStyle, ["a:firstRow", "a:tcTxStyle"]);
            if (rowTxtStyl !== undefined) {
                let local_fontClrPr = PPTXStyleUtils.getSolidFill(rowTxtStyl, undefined, undefined, warpObj);
                if (local_fontClrPr !== undefined) {
                    fontClrPr = local_fontClrPr;
                }
                let local_fontWeight = ((PPTXXmlUtils.getTextByPathList(rowTxtStyl, ["attrs", "b"]) == "on") ? "bold" : "");
                if (local_fontWeight !== "") {
                    fontWeight = local_fontWeight;
                }
            }
        }
        else if (i > 0 && tblStylAttrObj["isBandRowAttr"] == 1 && thisTblStyle !== undefined) {
            fillColor = "";
            row_borders = undefined;
            if ((i % 2) == 0 && thisTblStyle["a:band2H"] !== undefined) {
                let bgFillschemeClr = PPTXXmlUtils.getTextByPathList(thisTblStyle, ["a:band2H", "a:tcStyle", "a:fill", "a:solidFill"]);
                if (bgFillschemeClr !== undefined) {
                    let local_fillColor = PPTXStyleUtils.getSolidFill(bgFillschemeClr, undefined, undefined, warpObj);
                    if (local_fillColor !== "") {
                        fillColor = local_fillColor;
                    }
                }
                let borderStyl = PPTXXmlUtils.getTextByPathList(thisTblStyle, ["a:band2H", "a:tcStyle", "a:tcBdr"]);
                if (borderStyl !== undefined) {
                    let local_row_borders = PPTXStyleUtils.getTableBorders(borderStyl, warpObj);
                    if (local_row_borders != "") {
                        row_borders = local_row_borders;
                    }
                }
                let rowTxtStyl = PPTXXmlUtils.getTextByPathList(thisTblStyle, ["a:band2H", "a:tcTxStyle"]);
                if (rowTxtStyl !== undefined) {
                    let local_fontClrPr = PPTXStyleUtils.getSolidFill(rowTxtStyl, undefined, undefined, warpObj);
                    if (local_fontClrPr !== undefined) {
                        fontClrPr = local_fontClrPr;
                    }
                }
                let local_fontWeight = ((PPTXXmlUtils.getTextByPathList(rowTxtStyl, ["attrs", "b"]) == "on") ? "bold" : "");
                if (local_fontWeight !== "") {
                    fontWeight = local_fontWeight;
                }
            }
            if ((i % 2) != 0 && thisTblStyle["a:band1H"] !== undefined) {
                let bgFillschemeClr = PPTXXmlUtils.getTextByPathList(thisTblStyle, ["a:band1H", "a:tcStyle", "a:fill", "a:solidFill"]);
                if (bgFillschemeClr !== undefined) {
                    let local_fillColor = PPTXStyleUtils.getSolidFill(bgFillschemeClr, undefined, undefined, warpObj);
                    if (local_fillColor !== undefined) {
                        fillColor = local_fillColor;
                    }
                }
                let borderStyl = PPTXXmlUtils.getTextByPathList(thisTblStyle, ["a:band1H", "a:tcStyle", "a:tcBdr"]);
                if (borderStyl !== undefined) {
                    let local_row_borders = PPTXStyleUtils.getTableBorders(borderStyl, warpObj);
                    if (local_row_borders != "") {
                        row_borders = local_row_borders;
                    }
                }
                let rowTxtStyl = PPTXXmlUtils.getTextByPathList(thisTblStyle, ["a:band1H", "a:tcTxStyle"]);
                if (rowTxtStyl !== undefined) {
                    let local_fontClrPr = PPTXStyleUtils.getSolidFill(rowTxtStyl, undefined, undefined, warpObj);
                    if (local_fontClrPr !== undefined) {
                        fontClrPr = local_fontClrPr;
                    }
                    let local_fontWeight = ((PPTXXmlUtils.getTextByPathList(rowTxtStyl, ["attrs", "b"]) == "on") ? "bold" : "");
                    if (local_fontWeight != "") {
                        fontWeight = local_fontWeight;
                    }
                }
            }
        }
        if (i == (trNodes.length - 1) && tblStylAttrObj["isLstRowAttr"] == 1 && thisTblStyle !== undefined) {
            let bgFillschemeClr = PPTXXmlUtils.getTextByPathList(thisTblStyle, ["a:lastRow", "a:tcStyle", "a:fill", "a:solidFill"]);
            if (bgFillschemeClr !== undefined) {
                let local_fillColor = PPTXStyleUtils.getSolidFill(bgFillschemeClr, undefined, undefined, warpObj);
                if (local_fillColor !== undefined) {
                    fillColor = local_fillColor;
                }
            }
            let borderStyl = PPTXXmlUtils.getTextByPathList(thisTblStyle, ["a:lastRow", "a:tcStyle", "a:tcBdr"]);
            if (borderStyl !== undefined) {
                let local_row_borders = PPTXStyleUtils.getTableBorders(borderStyl, warpObj);
                if (local_row_borders != "") {
                    row_borders = local_row_borders;
                }
            }
            let rowTxtStyl = PPTXXmlUtils.getTextByPathList(thisTblStyle, ["a:lastRow", "a:tcTxStyle"]);
            if (rowTxtStyl !== undefined) {
                let local_fontClrPr = PPTXStyleUtils.getSolidFill(rowTxtStyl, undefined, undefined, warpObj);
                if (local_fontClrPr !== undefined) {
                    fontClrPr = local_fontClrPr;
                }
                let local_fontWeight = ((PPTXXmlUtils.getTextByPathList(rowTxtStyl, ["attrs", "b"]) == "on") ? "bold" : "");
                if (local_fontWeight !== "") {
                    fontWeight = local_fontWeight;
                }
            }
        }
        rowsStyl += ((row_borders !== undefined) ? row_borders : "");
        if (fontClrPr !== undefined && typeof fontClrPr === 'string') {
            let tableColorValue = fontClrPr;
            if (tableColorValue.length === 8) {
                let colorObj = tinycolor(tableColorValue);
                tableColorValue = colorObj.toRgbString();
            }
            else {
                tableColorValue = `#${tableColorValue}`;
            }
            rowsStyl += ` color: ${tableColorValue};`;
        }
        rowsStyl += ((fontWeight != "") ? ` font-weight:${fontWeight};` : "");
        if (fillColor !== undefined && fillColor != "" && typeof fillColor === 'string') {
            if (fillColor.length === 8) {
                let colorObj = tinycolor(fillColor);
                fillColor = colorObj.toRgbString();
            }
            else {
                fillColor = `#${fillColor}`;
            }
            rowsStyl += `background-color: ${fillColor};`;
        }
        tableHtml += `<tr style='${rowsStyl}'>`;
        let tcNodes = trNodes[i]["a:tc"];
        if (tcNodes !== undefined) {
            if (tcNodes.constructor === Array) {
                let j = 0;
                if (rowSpanAry.length == 0) {
                    rowSpanAry = Array.apply(null, Array(tcNodes.length)).map(() => { return 0; });
                }
                let totalColSpan = 0;
                while (j < tcNodes.length) {
                    if (rowSpanAry[j] == 0 && totalColSpan == 0) {
                        let a_sorce;
                        if (j == 0 && tblStylAttrObj["isFrstColAttr"] == 1) {
                            a_sorce = "a:firstCol";
                            if (tblStylAttrObj["isLstRowAttr"] == 1 && i == (trNodes.length - 1) &&
                                PPTXXmlUtils.getTextByPathList(thisTblStyle, ["a:seCell"]) !== undefined) {
                                a_sorce = "a:seCell";
                            }
                            else if (tblStylAttrObj["isFrstRowAttr"] == 1 && i == 0 &&
                                PPTXXmlUtils.getTextByPathList(thisTblStyle, ["a:neCell"]) !== undefined) {
                                a_sorce = "a:neCell";
                            }
                        }
                        else if ((j > 0 && tblStylAttrObj["isBandColAttr"] == 1) &&
                            !(tblStylAttrObj["isFrstColAttr"] == 1 && i == 0) &&
                            !(tblStylAttrObj["isLstRowAttr"] == 1 && i == (trNodes.length - 1)) &&
                            j != (tcNodes.length - 1)) {
                            if ((j % 2) != 0) {
                                let aBandNode = PPTXXmlUtils.getTextByPathList(thisTblStyle, ["a:band2V"]);
                                if (aBandNode === undefined) {
                                    aBandNode = PPTXXmlUtils.getTextByPathList(thisTblStyle, ["a:band1V"]);
                                    if (aBandNode !== undefined) {
                                        a_sorce = "a:band2V";
                                    }
                                }
                                else {
                                    a_sorce = "a:band2V";
                                }
                            }
                        }
                        if (j == (tcNodes.length - 1) && tblStylAttrObj["isLstColAttr"] == 1) {
                            a_sorce = "a:lastCol";
                            if (tblStylAttrObj["isLstRowAttr"] == 1 && i == (trNodes.length - 1) && PPTXXmlUtils.getTextByPathList(thisTblStyle, ["a:swCell"]) !== undefined) {
                                a_sorce = "a:swCell";
                            }
                            else if (tblStylAttrObj["isFrstRowAttr"] == 1 && i == 0 && PPTXXmlUtils.getTextByPathList(thisTblStyle, ["a:nwCell"]) !== undefined) {
                                a_sorce = "a:nwCell";
                            }
                        }
                        let cellParmAry = await getTableCellParams(tcNodes[j], getColsGrid, i, j, thisTblStyle, a_sorce, warpObj);
                        let text = cellParmAry[0];
                        let colStyl = cellParmAry[1];
                        let cssName = cellParmAry[2];
                        let rowSpan = cellParmAry[3];
                        let colSpan = cellParmAry[4];
                        if (rowSpan !== undefined) {
                            rowSpanAry[j] = parseInt(rowSpan) - 1;
                            tableHtml += `<td class='${cssName}' data-row='` + i + `,${j}' rowspan ='` +
                                parseInt(rowSpan) + `' style='${colStyl}'>` + text + "</td>";
                        }
                        else if (colSpan !== undefined) {
                            tableHtml += `<td class='${cssName}' data-row='` + i + `,${j}' colspan = '` +
                                parseInt(colSpan) + `' style='${colStyl}'>` + text + "</td>";
                            totalColSpan = parseInt(colSpan) - 1;
                        }
                        else {
                            tableHtml += `<td class='${cssName}' data-row='` + i + `,${j}' style = '` + colStyl + `'>${text}</td>`;
                        }
                    }
                    else {
                        if (rowSpanAry[j] != 0) {
                            rowSpanAry[j] -= 1;
                        }
                        if (totalColSpan != 0) {
                            totalColSpan--;
                        }
                    }
                    j++;
                }
            }
            else {
                let a_sorce;
                if (tblStylAttrObj["isFrstColAttr"] == 1 && !(tblStylAttrObj["isLstRowAttr"] == 1)) {
                    a_sorce = "a:firstCol";
                }
                else if ((tblStylAttrObj["isBandColAttr"] == 1) && !(tblStylAttrObj["isLstRowAttr"] == 1)) {
                    let aBandNode = PPTXXmlUtils.getTextByPathList(thisTblStyle, ["a:band2V"]);
                    if (aBandNode === undefined) {
                        aBandNode = PPTXXmlUtils.getTextByPathList(thisTblStyle, ["a:band1V"]);
                        if (aBandNode !== undefined) {
                            a_sorce = "a:band2V";
                        }
                    }
                    else {
                        a_sorce = "a:band2V";
                    }
                }
                if (tblStylAttrObj["isLstColAttr"] == 1 && !(tblStylAttrObj["isLstRowAttr"] == 1)) {
                    a_sorce = "a:lastCol";
                }
                let cellParmAry = await getTableCellParams(tcNodes, getColsGrid, i, undefined, thisTblStyle, a_sorce, warpObj);
                let text = cellParmAry[0];
                let colStyl = cellParmAry[1];
                let cssName = cellParmAry[2];
                let rowSpan = cellParmAry[3];
                if (rowSpan !== undefined) {
                    tableHtml += `<td  class='${cssName}' rowspan='` + parseInt(rowSpan) + `' style = '${colStyl}'>` + text + "</td>";
                }
                else {
                    tableHtml += `<td class='${cssName}' style='` + colStyl + `'>${text}</td>`;
                }
            }
        }
        tableHtml += "</tr>";
    }
    return tableHtml;
}
async function getTableCellParams(tcNodes, getColsGrid, row_idx, col_idx, thisTblStyle, cellSource, warpObj) {
    let rowSpan = PPTXXmlUtils.getTextByPathList(tcNodes, ["attrs", "rowSpan"]);
    let colSpan = PPTXXmlUtils.getTextByPathList(tcNodes, ["attrs", "gridSpan"]);
    PPTXXmlUtils.getTextByPathList(tcNodes, ["attrs", "vMerge"]);
    PPTXXmlUtils.getTextByPathList(tcNodes, ["attrs", "hMerge"]);
    let colStyl = "word-wrap: break-word;";
    let colWidth;
    let celFillColor = "";
    let colFontClrPr = "";
    let colFontWeight = "";
    let lin_bottm = "", lin_top = "", lin_left = "", lin_right = "";
    let colSapnInt = parseInt(colSpan);
    let total_col_width = 0;
    if (!isNaN(colSapnInt) && colSapnInt > 1) {
        for (let k = 0; k < colSapnInt; k++) {
            total_col_width += parseInt(PPTXXmlUtils.getTextByPathList(getColsGrid[col_idx + k], ["attrs", "w"]));
        }
    }
    else {
        total_col_width = PPTXXmlUtils.getTextByPathList((col_idx === undefined) ? getColsGrid : getColsGrid[col_idx], ["attrs", "w"]);
    }
    let text = await PPTXTextUtils.genTextBody(tcNodes["a:txBody"], tcNodes, undefined, undefined, "table", undefined, warpObj, total_col_width);
    if (total_col_width != 0) {
        colWidth = parseInt(String(total_col_width)) * SLIDE_FACTOR$1;
        colWidth = Math.round(colWidth * 100) / 100;
        colStyl += `width:${colWidth}px;`;
    }
    let cellAlign = "";
    let textBodyNode = tcNodes["a:txBody"];
    let prg_dir = "";
    if (textBodyNode !== undefined) {
        let pNodes = textBodyNode["a:p"];
        if (pNodes !== undefined) {
            if (Array.isArray(pNodes) && pNodes.length > 0) {
                prg_dir = PPTXStyleUtils.getPregraphDir(pNodes[0], textBodyNode, 0, "table", warpObj);
            }
            else {
                prg_dir = PPTXStyleUtils.getPregraphDir(pNodes, textBodyNode, 0, "table", warpObj);
            }
        }
    }
    let isRTL = (prg_dir == "pregraph-rtl");
    if (textBodyNode !== undefined) {
        let pNodes = textBodyNode["a:p"];
        if (pNodes !== undefined) {
            let firstP = Array.isArray(pNodes) ? pNodes[0] : pNodes;
            let horizontalAlign = PPTXStyleUtils.getHorizontalAlign(firstP, textBodyNode, 0, "table", prg_dir, warpObj);
            if (horizontalAlign === "h-right" || horizontalAlign === "h-right-rtl") {
                cellAlign = isRTL ? "text-align: left;" : "text-align: right;";
            }
            else if (horizontalAlign === "h-mid") {
                cellAlign = "text-align: center;";
            }
            else if (horizontalAlign === "h-left-rtl") {
                cellAlign = "text-align: right;";
            }
            else {
                cellAlign = isRTL ? "text-align: right;" : "text-align: left;";
            }
        }
    }
    if (cellAlign !== "") {
        colStyl += cellAlign;
    }
    lin_bottm = PPTXXmlUtils.getTextByPathList(tcNodes, ["a:tcPr", "a:lnB"]);
    if (lin_bottm === undefined && cellSource !== undefined) {
        if (cellSource !== undefined)
            lin_bottm = PPTXXmlUtils.getTextByPathList(thisTblStyle[cellSource], ["a:tcStyle", "a:tcBdr", "a:bottom", "a:ln"]);
        if (lin_bottm === undefined) {
            lin_bottm = PPTXXmlUtils.getTextByPathList(thisTblStyle, ["a:wholeTbl", "a:tcStyle", "a:tcBdr", "a:bottom", "a:ln"]);
        }
    }
    lin_top = PPTXXmlUtils.getTextByPathList(tcNodes, ["a:tcPr", "a:lnT"]);
    if (lin_top === undefined) {
        if (cellSource !== undefined)
            lin_top = PPTXXmlUtils.getTextByPathList(thisTblStyle[cellSource], ["a:tcStyle", "a:tcBdr", "a:top", "a:ln"]);
        if (lin_top === undefined) {
            lin_top = PPTXXmlUtils.getTextByPathList(thisTblStyle, ["a:wholeTbl", "a:tcStyle", "a:tcBdr", "a:top", "a:ln"]);
        }
    }
    lin_left = PPTXXmlUtils.getTextByPathList(tcNodes, ["a:tcPr", "a:lnL"]);
    if (lin_left === undefined) {
        if (cellSource !== undefined)
            lin_left = PPTXXmlUtils.getTextByPathList(thisTblStyle[cellSource], ["a:tcStyle", "a:tcBdr", "a:left", "a:ln"]);
        if (lin_left === undefined) {
            lin_left = PPTXXmlUtils.getTextByPathList(thisTblStyle, ["a:wholeTbl", "a:tcStyle", "a:tcBdr", "a:left", "a:ln"]);
        }
    }
    lin_right = PPTXXmlUtils.getTextByPathList(tcNodes, ["a:tcPr", "a:lnR"]);
    if (lin_right === undefined) {
        if (cellSource !== undefined)
            lin_right = PPTXXmlUtils.getTextByPathList(thisTblStyle[cellSource], ["a:tcStyle", "a:tcBdr", "a:right", "a:ln"]);
        if (lin_right === undefined) {
            lin_right = PPTXXmlUtils.getTextByPathList(thisTblStyle, ["a:wholeTbl", "a:tcStyle", "a:tcBdr", "a:right", "a:ln"]);
        }
    }
    PPTXXmlUtils.getTextByPathList(tcNodes, ["a:tcPr", "a:lnBlToTr"]);
    PPTXXmlUtils.getTextByPathList(tcNodes, ["a:tcPr", "a:InTlToBr"]);
    if (lin_bottm !== undefined && lin_bottm != "") {
        let bottom_line_border = PPTXStyleUtils.getBorder(lin_bottm, undefined, false, "", warpObj);
        if (bottom_line_border != "") {
            colStyl += `border-bottom:${bottom_line_border};`;
        }
    }
    if (lin_top !== undefined && lin_top != "") {
        let top_line_border = PPTXStyleUtils.getBorder(lin_top, undefined, false, "", warpObj);
        if (top_line_border != "") {
            colStyl += `border-top: ${top_line_border};`;
        }
    }
    if (lin_left !== undefined && lin_left != "") {
        let left_line_border = PPTXStyleUtils.getBorder(lin_left, undefined, false, "", warpObj);
        if (left_line_border != "") {
            colStyl += `border-left: ${left_line_border};`;
        }
    }
    if (lin_right !== undefined && lin_right != "") {
        let right_line_border = PPTXStyleUtils.getBorder(lin_right, undefined, false, "", warpObj);
        if (right_line_border != "") {
            colStyl += `border-right:${right_line_border};`;
        }
    }
    let getCelFill = PPTXXmlUtils.getTextByPathList(tcNodes, ["a:tcPr"]);
    if (getCelFill !== undefined && getCelFill != "") {
        let cellObj = {
            "p:spPr": getCelFill
        };
        celFillColor = await PPTXStyleUtils.getShapeFill(cellObj, undefined, false, warpObj, "slide");
    }
    if (celFillColor == "" || celFillColor == "background-color: inherit;") {
        let bgFillschemeClr;
        if (cellSource !== undefined)
            bgFillschemeClr = PPTXXmlUtils.getTextByPathList(thisTblStyle, [cellSource, "a:tcStyle", "a:fill", "a:solidFill"]);
        if (bgFillschemeClr !== undefined) {
            let local_fillColor = PPTXStyleUtils.getSolidFill(bgFillschemeClr, undefined, undefined, warpObj);
            if (local_fillColor !== undefined) {
                celFillColor = ` background-color: #${local_fillColor};`;
            }
        }
    }
    let cssName = "";
    if (celFillColor !== undefined && celFillColor != "") {
        if (celFillColor in warpObj.styleTable) {
            cssName = warpObj.styleTable[celFillColor]["name"];
        }
        else {
            cssName = `_tbl_cell_css_${(Object.keys(warpObj.styleTable).length + 1)}`;
            warpObj.styleTable[celFillColor] = {
                "name": cssName,
                "text": celFillColor
            };
        }
    }
    let rowTxtStyl;
    if (cellSource !== undefined) {
        rowTxtStyl = PPTXXmlUtils.getTextByPathList(thisTblStyle, [cellSource, "a:tcTxStyle"]);
    }
    if (rowTxtStyl !== undefined) {
        let local_fontClrPr = PPTXStyleUtils.getSolidFill(rowTxtStyl, undefined, undefined, warpObj);
        if (local_fontClrPr !== undefined) {
            colFontClrPr = local_fontClrPr;
        }
        let local_fontWeight = ((PPTXXmlUtils.getTextByPathList(rowTxtStyl, ["attrs", "b"]) == "on") ? "bold" : "");
        if (local_fontWeight !== "") {
            colFontWeight = local_fontWeight;
        }
    }
    colStyl += ((colFontClrPr !== "" && typeof colFontClrPr === 'string') ?
        ((colFontClrPr.length === 8) ?
            (() => {
                let colorObj = tinycolor(colFontClrPr);
                return `color: ${colorObj.toRgbString()};`;
            })() :
            `color: #${colFontClrPr};`) : "");
    colStyl += ((colFontWeight != "") ? ` font-weight:${colFontWeight};` : "");
    return [text, colStyl, cssName, rowSpan, colSpan];
}
const PPTXTextUtils = {
    genTextBody,
    genBuChar,
    getHtmlBullet,
    getDingbatToUnicode,
    genSpanElement,
    genTable,
    getTableCellParams,
    alphaNumeric,
    archaicNumbers,
    romanize,
    getNumTypeNum,
};

function polarToCartesian(cx, cy, w, h, angleInDegrees) {
    const angleInRadians = (angleInDegrees - 90) * Math.PI / 180.0;
    function fmt(num) {
        return parseFloat(num.toFixed(2));
    }
    return {
        x: fmt(cx + (w / 2) * Math.cos(angleInRadians)),
        y: fmt(cy + (h / 2) * Math.sin(angleInRadians))
    };
}
function shapeArc(cx, cy, w, h, startAngle, endAngle, clockwise) {
    const start = polarToCartesian(cx, cy, w, h, endAngle);
    const end = polarToCartesian(cx, cy, w, h, startAngle);
    const largeArcFlag = endAngle - startAngle <= 180 ? "0" : "1";
    function fmt(num) {
        return parseFloat(num.toFixed(2));
    }
    const d = [
        "M", start.x, start.y,
        "A", fmt(w), fmt(h), 0, largeArcFlag, clockwise ? "0" : "1", end.x, end.y
    ].join(" ");
    return d;
}
function shapeArcAlt(cX, cY, rX, rY, stAng, endAng, isClose) {
    let dData;
    let angle = stAng;
    function fmt(num) {
        return parseFloat(num.toFixed(2));
    }
    if (endAng >= stAng) {
        while (angle <= endAng) {
            const radians = angle * (Math.PI / 180);
            const x = cX + Math.cos(radians) * rX;
            const y = cY + Math.sin(radians) * rY;
            if (angle == stAng) {
                dData = ` M${fmt(x)} ${fmt(y)}`;
            }
            dData += ` L${fmt(x)} ${fmt(y)}`;
            angle++;
        }
    }
    else {
        while (angle > endAng) {
            const radians = angle * (Math.PI / 180);
            const x = cX + Math.cos(radians) * rX;
            const y = cY + Math.sin(radians) * rY;
            if (angle == stAng) {
                dData = ` M ${fmt(x)} ${fmt(y)}`;
            }
            dData += ` L ${fmt(x)} ${fmt(y)}`;
            angle--;
        }
    }
    dData += (isClose ? " z" : "");
    return dData;
}
function shapeSnipRoundRect(w, h, sAdj1_val, sAdj2_val, shpTyp, adjTyp) {
    let d = "";
    let sAdj1 = 0;
    let sAdj2 = 0;
    if (shpTyp == "round") {
        sAdj1 = w * sAdj1_val;
        if (adjTyp == "cornrAll") {
            d = `M0,${sAdj1} Q0,0 ${sAdj1},0 L${(w - sAdj1)},0 Q${w},0 ${w},${sAdj1} L${w},${(h - sAdj1)} Q${w},${h} ${(w - sAdj1)},${h} L${sAdj1},${h} Q0,${h} 0,${(h - sAdj1)} z`;
        }
        else if (adjTyp == "cornr1") {
            d = `M0,0 L${(w - sAdj1)},0 Q${w},0 ${w},${sAdj1} L${w},${h} L0,${h} z`;
        }
        else if (adjTyp == "diag") {
            sAdj2 = h * sAdj2_val;
            d = `M0,0 L${(w - sAdj1)},0 Q${w},0 ${w},${sAdj1} L${w},${(h - sAdj2)} Q${w},${h} ${(w - sAdj2)},${h} L${sAdj1},${h} Q0,${h} 0,${(h - sAdj1)} L0,${sAdj2} Q0,0 ${sAdj2},0 z`;
        }
        else if (adjTyp == "cornr2") {
            sAdj2 = w * sAdj2_val;
            d = `M0,0 L${(w - sAdj1)},0 Q${w},0 ${w},${sAdj1} L${w},${(h - sAdj2)} Q${w},${h} ${(w - sAdj2)},${h} L0,${h} z`;
        }
    }
    else if (shpTyp == "snip") {
        sAdj1 = w * sAdj1_val;
        if (adjTyp == "cornr1") {
            d = `M${sAdj1},0 L${w},0 L${w},${h} L0,${h} L0,${sAdj1} z`;
        }
        else if (adjTyp == "diag") {
            sAdj2 = h * sAdj2_val;
            d = `M${sAdj1},0 L${w},0 L${w},${(h - sAdj2)} L${sAdj2},${h} L0,${h} L0,${sAdj1} z`;
        }
        else if (adjTyp == "cornr2") {
            sAdj2 = w * sAdj2_val;
            d = `M${sAdj1},0 L${w},0 L${w},${(h - sAdj2)} L${(w - sAdj2)},${h} L0,${h} z`;
        }
    }
    return d;
}
function shapeSnipRoundRectAlt(w, h, adj1, adj2, shapeType, adjType) {
    let adjA, adjB, adjC, adjD;
    if (adjType == "cornr1") {
        adjA = 0;
        adjB = 0;
        adjC = 0;
        adjD = adj1;
    }
    else if (adjType == "cornr2") {
        adjA = adj1;
        adjB = adj2;
        adjC = adj2;
        adjD = adj1;
    }
    else if (adjType == "cornrAll") {
        adjA = adj1;
        adjB = adj1;
        adjC = adj1;
        adjD = adj1;
    }
    else if (adjType == "diag") {
        adjA = adj1;
        adjB = adj2;
        adjC = adj1;
        adjD = adj2;
    }
    let d;
    if (shapeType == "round") {
        d = `M0,${(h / 2 + (1 - adjB) * (h / 2))} Q${0},${h} ${adjB * (w / 2)},${h} L${(w / 2 + (1 - adjC) * (w / 2))},${h} Q${w},${h} ${w},${(h / 2 + (h / 2) * (1 - adjC))}L${w},${(h / 2) * adjD} Q${w},${0} ${(w / 2 + (w / 2) * (1 - adjD))},0 L${(w / 2) * adjA},0 Q${0},${0} 0,${(h / 2) * (adjA)} z`;
    }
    else if (shapeType == "snip") {
        d = `M0,${adjA * (h / 2)} L0,${(h / 2 + (h / 2) * (1 - adjB))}L${adjB * (w / 2)},${h} L${(w / 2 + (w / 2) * (1 - adjC))},${h}L${w},${(h / 2 + (h / 2) * (1 - adjC))} L${w},${adjD * (h / 2)}L${(w / 2 + (w / 2) * (1 - adjD))},0 L${((w / 2) * adjA)},0 z`;
    }
    return d;
}
function shapePie(H, w, adj1, adj2, isClose) {
    const pieVal = parseInt(adj2);
    const piAngle = parseInt(adj1);
    let size = parseInt(H), radius = (size / 2), value = pieVal - piAngle;
    if (value < 0) {
        value = 360 + value;
    }
    value = Math.min(Math.max(value, 0), 360);
    const x = Math.cos((2 * Math.PI) / (360 / value));
    const y = Math.sin((2 * Math.PI) / (360 / value));
    let longArc, d, rot;
    if (isClose) {
        longArc = (value <= 180) ? 0 : 1;
        d = `M${radius},${radius} L${radius},${0} A${radius},${radius} 0 ${longArc},1 ${(radius + y * radius)},${(radius - x * radius)} z`;
        rot = `rotate(${(piAngle - 270)}, ${radius}, ${radius})`;
    }
    else {
        longArc = (value <= 180) ? 0 : 1;
        const radius1 = radius;
        const radius2 = w / 2;
        d = `M${radius1},${0} A${radius2},${radius1} 0 ${longArc},1 ${(radius2 + y * radius2)},${(radius1 - x * radius1)}`;
        rot = `rotate(${(piAngle + 90)}, ${radius}, ${radius})`;
    }
    return [d, rot];
}
function shapeGear(w, h, points) {
    const innerRadius = h;
    const outerRadius = 1.5 * innerRadius;
    const cx = outerRadius;
    const cy = outerRadius;
    const notches = points;
    const radiusO = outerRadius;
    const radiusI = innerRadius;
    const taperO = 50;
    const taperI = 35;
    const pi2 = 2 * Math.PI;
    const angle = pi2 / (notches * 2);
    const taperAI = angle * taperI * 0.005;
    const taperAO = angle * taperO * 0.005;
    let a = angle;
    let toggle = false;
    let d = ` M${(cx + radiusO * Math.cos(taperAO))} ${(cy + radiusO * Math.sin(taperAO))}`;
    for (; a <= pi2 + angle; a += angle) {
        if (toggle) {
            d += ` L${(cx + radiusI * Math.cos(a - taperAI))},${(cy + radiusI * Math.sin(a - taperAI))}`;
            d += ` L${(cx + radiusO * Math.cos(a + taperAO))},${(cy + radiusO * Math.sin(a + taperAO))}`;
        }
        else {
            d += ` L${(cx + radiusO * Math.cos(a - taperAO))},${(cy + radiusO * Math.sin(a - taperAO))}`;
            d += ` L${(cx + radiusI * Math.cos(a + taperAI))},${(cy + radiusI * Math.sin(a + taperAI))}`;
        }
        toggle = !toggle;
    }
    d += " ";
    return d;
}

function renderCustomShape(custShapType, w, h, imgFillFlg, grndFillFlg, fillColor, border, shpId, shapeArcFn) {
    const pathLstNode = PPTXXmlUtils.getTextByPathList(custShapType, ["a:pathLst"]);
    const pathNodes = PPTXXmlUtils.getTextByPathList(pathLstNode, ["a:path"]);
    let maxX = 0;
    let maxY = 0;
    if (pathNodes && pathNodes["attrs"]) {
        maxX = parseInt(pathNodes["attrs"]["w"]) || 0;
        maxY = parseInt(pathNodes["attrs"]["h"]) || 0;
    }
    if (maxX <= 0)
        maxX = 1;
    if (maxY <= 0)
        maxY = 1;
    let cX = (1 / maxX) * w;
    let cY = (1 / maxY) * h;
    let moveToNode = PPTXXmlUtils.getTextByPathList(pathNodes, ["a:moveTo"]);
    moveToNode.length;
    const lnToNodes = pathNodes["a:lnTo"];
    let cubicBezToNodes = pathNodes["a:cubicBezTo"];
    const arcToNodes = pathNodes["a:arcTo"];
    let closeNode = PPTXXmlUtils.getTextByPathList(pathNodes, ["a:close"]);
    if (!Array.isArray(moveToNode)) {
        moveToNode = [moveToNode];
    }
    const multiSapeAry = [];
    if (moveToNode.length > 0) {
        Object.keys(moveToNode).forEach((key) => {
            var moveToPtNode = moveToNode[key]["a:pt"];
            if (moveToPtNode !== undefined) {
                Object.keys(moveToPtNode).forEach((key2) => {
                    var ptObj = {};
                    var moveToNoPt = moveToPtNode[key2];
                    var spX = moveToNoPt["x"];
                    var spY = moveToNoPt["y"];
                    var ptOrdr = moveToNoPt["order"];
                    ptObj.type = "movto";
                    ptObj.order = ptOrdr;
                    ptObj.x = spX;
                    ptObj.y = spY;
                    multiSapeAry.push(ptObj);
                });
            }
        });
        if (lnToNodes !== undefined) {
            Object.keys(lnToNodes).forEach((key) => {
                var lnToPtNode = lnToNodes[key]["a:pt"];
                if (lnToPtNode !== undefined) {
                    Object.keys(lnToPtNode).forEach((key2) => {
                        var ptObj = {};
                        var lnToNoPt = lnToPtNode[key2];
                        var ptX = lnToNoPt["x"];
                        var ptY = lnToNoPt["y"];
                        var ptOrdr = lnToNoPt["order"];
                        ptObj.type = "lnto";
                        ptObj.order = ptOrdr;
                        ptObj.x = ptX;
                        ptObj.y = ptY;
                        multiSapeAry.push(ptObj);
                    });
                }
            });
        }
        if (cubicBezToNodes !== undefined) {
            const cubicBezToPtNodesAry = [];
            if (!Array.isArray(cubicBezToNodes)) {
                cubicBezToNodes = [cubicBezToNodes];
            }
            Object.keys(cubicBezToNodes).forEach((key) => {
                cubicBezToPtNodesAry.push(cubicBezToNodes[key]["a:pt"]);
            });
            cubicBezToPtNodesAry.forEach((key2) => {
                var nodeObj = {};
                nodeObj.type = "cubicBezTo";
                nodeObj.order = key2[0]["attrs"]["order"];
                var pts_ary = [];
                key2.forEach((pt) => {
                    var pt_obj = {
                        x: pt["attrs"]["x"],
                        y: pt["attrs"]["y"]
                    };
                    pts_ary.push(pt_obj);
                });
                nodeObj.cubBzPt = pts_ary;
                multiSapeAry.push(nodeObj);
            });
        }
        let quadBezToNodes = pathNodes["a:quadBezTo"];
        if (quadBezToNodes !== undefined) {
            const quadBezToPtNodesAry = [];
            if (!Array.isArray(quadBezToNodes)) {
                quadBezToNodes = [quadBezToNodes];
            }
            Object.keys(quadBezToNodes).forEach((key) => {
                quadBezToPtNodesAry.push(quadBezToNodes[key]["a:pt"]);
            });
            quadBezToPtNodesAry.forEach((key2) => {
                var nodeObj = {};
                nodeObj.type = "quadBezTo";
                nodeObj.order = key2[0]["attrs"]["order"];
                var pts_ary = [];
                key2.forEach((pt) => {
                    var pt_obj = {
                        x: pt["attrs"]["x"],
                        y: pt["attrs"]["y"]
                    };
                    pts_ary.push(pt_obj);
                });
                nodeObj.quadBzPt = pts_ary;
                multiSapeAry.push(nodeObj);
            });
        }
        if (arcToNodes !== undefined) {
            const arcToNodesAttrs = arcToNodes["attrs"];
            const arcOrder = arcToNodesAttrs["order"];
            const hR = arcToNodesAttrs["hR"];
            const wR = arcToNodesAttrs["wR"];
            let stAng = arcToNodesAttrs["stAng"];
            let swAng = arcToNodesAttrs["swAng"];
            let shftX = 0;
            let shftY = 0;
            const arcToPtNode = PPTXXmlUtils.getTextByPathList(arcToNodes, ["a:pt", "attrs"]);
            if (arcToPtNode !== undefined) {
                shftX = arcToPtNode["x"];
                shftY = arcToPtNode["y"];
            }
            var ptObj = {};
            ptObj.type = "arcTo";
            ptObj.order = arcOrder;
            ptObj.hR = hR;
            ptObj.wR = wR;
            ptObj.stAng = stAng;
            ptObj.swAng = swAng;
            ptObj.shftX = shftX;
            ptObj.shftY = shftY;
            multiSapeAry.push(ptObj);
        }
        if (closeNode !== undefined) {
            if (!Array.isArray(closeNode)) {
                closeNode = [closeNode];
            }
            Object.keys(closeNode).forEach((key) => {
                var clsAttrs = closeNode[key]["attrs"];
                var clsOrder = clsAttrs["order"];
                var ptObj = {};
                ptObj.type = "close";
                ptObj.order = clsOrder;
                multiSapeAry.push(ptObj);
            });
        }
        multiSapeAry.sort((a, b) => {
            return a.order - b.order;
        });
        let k = 0;
        if (isNaN(cX))
            cX = 0;
        if (isNaN(cY))
            cY = 0;
        let d = "";
        while (k < multiSapeAry.length) {
            if (multiSapeAry[k].type == "movto") {
                const xVal = parseInt(multiSapeAry[k].x) || 0;
                const yVal = parseInt(multiSapeAry[k].y) || 0;
                if (isNaN(cX))
                    cX = 0;
                if (isNaN(cY))
                    cY = 0;
                var spX = xVal * cX;
                var spY = yVal * cY;
                d += ` M${spX},${spY}`;
            }
            else if (multiSapeAry[k].type == "lnto") {
                const xVal = parseInt(multiSapeAry[k].x) || 0;
                const yVal = parseInt(multiSapeAry[k].y) || 0;
                if (isNaN(cX))
                    cX = 0;
                if (isNaN(cY))
                    cY = 0;
                const Lx = xVal * cX;
                const Ly = yVal * cY;
                d += ` L${Lx},${Ly}`;
            }
            else if (multiSapeAry[k].type == "cubicBezTo") {
                if (isNaN(cX))
                    cX = 0;
                if (isNaN(cY))
                    cY = 0;
                const Cx1 = (parseInt(multiSapeAry[k].cubBzPt[0].x) || 0) * cX;
                const Cy1 = (parseInt(multiSapeAry[k].cubBzPt[0].y) || 0) * cY;
                const Cx2 = (parseInt(multiSapeAry[k].cubBzPt[1].x) || 0) * cX;
                const Cy2 = (parseInt(multiSapeAry[k].cubBzPt[1].y) || 0) * cY;
                const Cx3 = (parseInt(multiSapeAry[k].cubBzPt[2].x) || 0) * cX;
                const Cy3 = (parseInt(multiSapeAry[k].cubBzPt[2].y) || 0) * cY;
                d += ` C${Cx1},${Cy1} ${Cx2},${Cy2} ${Cx3},${Cy3}`;
            }
            else if (multiSapeAry[k].type == "arcTo") {
                if (isNaN(cX))
                    cX = 0;
                if (isNaN(cY))
                    cY = 0;
                const hR = (parseInt(multiSapeAry[k].hR) || 0) * cX;
                const wR = (parseInt(multiSapeAry[k].wR) || 0) * cY;
                let stAng = (parseInt(multiSapeAry[k].stAng) || 0) / 60000;
                let swAng = (parseInt(multiSapeAry[k].swAng) || 0) / 60000;
                if (isNaN(stAng))
                    stAng = 0;
                if (isNaN(swAng))
                    swAng = 0;
                const endAng = stAng + swAng;
                if (!isNaN(hR) && !isNaN(wR) && !isNaN(stAng) && !isNaN(swAng)) {
                    d += shapeArcFn(wR, hR, wR, hR, stAng, endAng, false);
                }
            }
            else if (multiSapeAry[k].type == "quadBezTo") {
                var quadBzPt = multiSapeAry[k].quadBzPt;
                if (quadBzPt && quadBzPt.length >= 2) {
                    const ctrlX = quadBzPt[0].x * cX;
                    const ctrlY = quadBzPt[0].y * cY;
                    const endX = quadBzPt[1].x * cX;
                    const endY = quadBzPt[1].y * cY;
                    d += `Q${ctrlX},${ctrlY} ${endX},${endY}`;
                }
            }
            else if (multiSapeAry[k].type == "close") {
                d += "z";
            }
            k++;
        }
        return `<path d='${d}' fill='${(!imgFillFlg ? (grndFillFlg ? "url(#linGrd_" + shpId + ")" : fillColor) : "url(#imgPtrn_" + shpId + ")")}' stroke='${((border === undefined) ? "" : border.color)}' stroke-width='${((border === undefined) ? "" : border.width)}' stroke-dasharray='${((border === undefined) ? "" : border.strokeDasharray)}' />`;
    }
    return "";
}

const SLIDE_FACTOR = 0.0001;
function renderStar(shapType, w, h, imgFillFlg, grndFillFlg, fillColor, border, shpId, shapeArcAlt, node) {
    let result = '';
    const hc = w / 2, vc = h / 2, wd2 = w / 2, hd2 = h / 2;
    const fill = !imgFillFlg ? (grndFillFlg ? `url(#linGrd_${shpId})` : fillColor) : `url(#imgPtrn_${shpId})`;
    switch (shapType) {
        case "star4": {
            const adj = getAdjValue(node, "adj", 19098);
            const cnstVal1 = 50000 * SLIDE_FACTOR;
            const a = clamp(adj, 0, cnstVal1);
            const iwd2 = wd2 * a / cnstVal1;
            const ihd2 = hd2 * a / cnstVal1;
            const sdx = iwd2 * Math.cos(0.7853981634);
            const sdy = ihd2 * Math.sin(0.7853981634);
            const sx1 = hc - sdx;
            const sx2 = hc + sdx;
            const sy1 = vc - sdy;
            const sy2 = vc + sdy;
            const d = `M0,${vc} L${sx1},${sy1} L${hc},0 L${sx2},${sy1} L${w},${vc} L${sx2},${sy2} L${hc},${h} L${sx1},${sy2} z`;
            result += `<path d='${d}' fill='${fill}' stroke='${border.color}' stroke-width='${border.width}' stroke-dasharray='${border.strokeDasharray}' />`;
            break;
        }
        case "star5": {
            const adj = getAdjValue(node, "adj", 19098);
            const hf = getAdjValue(node, "hf", 105146);
            const vf = getAdjValue(node, "vf", 110557);
            const maxAdj = 50000 * SLIDE_FACTOR;
            const cnstVal1 = 100000 * SLIDE_FACTOR;
            const a = clamp(adj, 0, maxAdj);
            const swd2 = wd2 * hf / cnstVal1;
            const shd2 = hd2 * vf / cnstVal1;
            const svc = vc * vf / cnstVal1;
            const dx1 = swd2 * Math.cos(0.31415926536);
            const dx2 = swd2 * Math.cos(5.3407075111);
            const dy1 = shd2 * Math.sin(0.31415926536);
            const dy2 = shd2 * Math.sin(5.3407075111);
            const x1 = hc - dx1;
            const x2 = hc - dx2;
            const x3 = hc + dx2;
            const x4 = hc + dx1;
            const y1 = svc - dy1;
            const y2 = svc - dy2;
            const iwd2 = swd2 * a / maxAdj;
            const ihd2 = shd2 * a / maxAdj;
            const sdx1 = iwd2 * Math.cos(5.9690260418);
            const sdx2 = iwd2 * Math.cos(0.94247779608);
            const sdy1 = ihd2 * Math.sin(0.94247779608);
            const sdy2 = ihd2 * Math.sin(5.9690260418);
            const sx1 = hc - sdx1;
            const sx2 = hc - sdx2;
            const sx3 = hc + sdx2;
            const sx4 = hc + sdx1;
            const sy1 = svc - sdy1;
            const sy2 = svc - sdy2;
            const sy3 = svc + ihd2;
            const d = `M${x1},${y1} L${sx2},${sy1} L${hc},${0} L${sx3},${sy1} L${x4},${y1} L${sx4},${sy2} L${x3},${y2} L${hc},${sy3} L${x2},${y2} L${sx1},${sy2} z`;
            result += `<path d='${d}' fill='${fill}' stroke='${border.color}' stroke-width='${border.width}' stroke-dasharray='${border.strokeDasharray}' />`;
            break;
        }
        case "star6": {
            const adj = getAdjValue(node, "adj", 28868);
            const hf = getAdjValue(node, "hf", 115470);
            const maxAdj = 50000 * SLIDE_FACTOR;
            const cnstVal1 = 100000 * SLIDE_FACTOR;
            const hd4 = h / 4;
            const a = clamp(adj, 0, maxAdj);
            const swd2 = wd2 * hf / cnstVal1;
            const dx1 = swd2 * Math.cos(0.5235987756);
            const x1 = hc - dx1;
            const x2 = hc + dx1;
            const y2 = vc + hd4;
            const iwd2 = swd2 * a / maxAdj;
            const ihd2 = hd2 * a / maxAdj;
            const sdx2 = iwd2 / 2;
            const sx1 = hc - iwd2;
            const sx2 = hc - sdx2;
            const sx3 = hc + sdx2;
            const sx4 = hc + iwd2;
            const sdy1 = ihd2 * Math.sin(1.0471975512);
            const sy1 = vc - sdy1;
            const sy2 = vc + sdy1;
            const d = `M${x1},${hd4} L${sx2},${sy1} L${hc},0 L${sx3},${sy1} L${x2},${hd4} L${sx4},${vc} L${x2},${y2} L${sx3},${sy2} L${hc},${h} L${sx2},${sy2} L${x1},${y2} L${sx1},${vc} z`;
            result += `<path d='${d}' fill='${fill}' stroke='${border.color}' stroke-width='${border.width}' stroke-dasharray='${border.strokeDasharray}' />`;
            break;
        }
        case "star7": {
            const adj = getAdjValue(node, "adj", 34601);
            const hf = getAdjValue(node, "hf", 102572);
            const vf = getAdjValue(node, "vf", 105210);
            const maxAdj = 50000 * SLIDE_FACTOR;
            const cnstVal1 = 100000 * SLIDE_FACTOR;
            const a = clamp(adj, 0, maxAdj);
            const swd2 = wd2 * hf / cnstVal1;
            const shd2 = hd2 * vf / cnstVal1;
            const svc = vc * vf / cnstVal1;
            const dx1 = swd2 * 97493 / 100000;
            const dx2 = swd2 * 78183 / 100000;
            const dx3 = swd2 * 43388 / 100000;
            const dy1 = shd2 * 62349 / 100000;
            const dy2 = shd2 * 22252 / 100000;
            const dy3 = shd2 * 90097 / 100000;
            const x1 = hc - dx1;
            const x2 = hc - dx2;
            const x3 = hc - dx3;
            const x4 = hc + dx3;
            const x5 = hc + dx2;
            const x6 = hc + dx1;
            const y1 = svc - dy1;
            const y2 = svc + dy2;
            const y3 = svc + dy3;
            const iwd2 = swd2 * a / maxAdj;
            const ihd2 = shd2 * a / maxAdj;
            const sdx1 = iwd2 * 97493 / 100000;
            const sdx2 = iwd2 * 78183 / 100000;
            const sdx3 = iwd2 * 43388 / 100000;
            const sx1 = hc - sdx1;
            const sx2 = hc - sdx2;
            const sx3 = hc - sdx3;
            const sx4 = hc + sdx3;
            const sx5 = hc + sdx2;
            const sx6 = hc + sdx1;
            const sdy1 = ihd2 * 90097 / 100000;
            const sdy2 = ihd2 * 22252 / 100000;
            const sdy3 = ihd2 * 62349 / 100000;
            const sy1 = svc - sdy1;
            const sy2 = svc - sdy2;
            const sy3 = svc + sdy3;
            const sy4 = svc + ihd2;
            const d = `M${x1},${y2} L${sx1},${sy2} L${x2},${y1} L${sx3},${sy1} L${hc},0 L${sx4},${sy1} L${x5},${y1} L${sx6},${sy2} L${x6},${y2} L${sx5},${sy3} L${x4},${y3} L${hc},${sy4} L${x3},${y3} L${sx2},${sy3} z`;
            result += `<path d='${d}' fill='${fill}' stroke='${border.color}' stroke-width='${border.width}' stroke-dasharray='${border.strokeDasharray}' />`;
            break;
        }
        case "star8": {
            const adj = getAdjValue(node, "adj", 37500);
            const maxAdj = 50000 * SLIDE_FACTOR;
            const a = clamp(adj, 0, maxAdj);
            const dx1 = wd2 * Math.cos(0.7853981634);
            const x1 = hc - dx1;
            const x2 = hc + dx1;
            const dy1 = hd2 * Math.sin(0.7853981634);
            const y1 = vc - dy1;
            const y2 = vc + dy1;
            const iwd2 = wd2 * a / maxAdj;
            const ihd2 = hd2 * a / maxAdj;
            const sdx1 = iwd2 * 92388 / 100000;
            const sdx2 = iwd2 * 38268 / 100000;
            const sdy1 = ihd2 * 92388 / 100000;
            const sdy2 = ihd2 * 38268 / 100000;
            const sx1 = hc - sdx1;
            const sx2 = hc - sdx2;
            const sx3 = hc + sdx2;
            const sx4 = hc + sdx1;
            const sy1 = vc - sdy1;
            const sy2 = vc - sdy2;
            const sy3 = vc + sdy2;
            const sy4 = vc + sdy1;
            const d = `M0,${vc} L${sx1},${sy2} L${x1},${y1} L${sx2},${sy1} L${hc},0 L${sx3},${sy1} L${x2},${y1} L${sx4},${sy2} L${w},${vc} L${sx4},${sy3} L${x2},${y2} L${sx3},${sy4} L${hc},${h} L${sx2},${sy4} L${x1},${y2} L${sx1},${sy3} z`;
            result += `<path d='${d}' fill='${fill}' stroke='${border.color}' stroke-width='${border.width}' stroke-dasharray='${border.strokeDasharray}' />`;
            break;
        }
        case "star10": {
            const adj = getAdjValue(node, "adj", 42533);
            const hf = getAdjValue(node, "hf", 105146);
            const maxAdj = 50000 * SLIDE_FACTOR;
            const cnstVal1 = 100000 * SLIDE_FACTOR;
            const a = clamp(adj, 0, maxAdj);
            const swd2 = wd2 * hf / cnstVal1;
            const dx1 = swd2 * 95106 / 100000;
            const dx2 = swd2 * 58779 / 100000;
            const x1 = hc - dx1;
            const x2 = hc - dx2;
            const x3 = hc + dx2;
            const x4 = hc + dx1;
            const dy1 = hd2 * 80902 / 100000;
            const dy2 = hd2 * 30902 / 100000;
            const y1 = vc - dy1;
            const y2 = vc - dy2;
            const y3 = vc + dy2;
            const y4 = vc + dy1;
            const iwd2 = swd2 * a / maxAdj;
            const ihd2 = hd2 * a / maxAdj;
            const sdx1 = iwd2 * 80902 / 100000;
            const sdx2 = iwd2 * 30902 / 100000;
            const sdy1 = ihd2 * 95106 / 100000;
            const sdy2 = ihd2 * 58779 / 100000;
            const sx1 = hc - iwd2;
            const sx2 = hc - sdx1;
            const sx3 = hc - sdx2;
            const sx4 = hc + sdx2;
            const sx5 = hc + sdx1;
            const sx6 = hc + iwd2;
            const sy1 = vc - sdy1;
            const sy2 = vc - sdy2;
            const sy3 = vc + sdy2;
            const sy4 = vc + sdy1;
            const d = `M${x1},${y2} L${sx2},${sy2} L${x2},${y1} L${sx3},${sy1} L${hc},0 L${sx4},${sy1} L${x3},${y1} L${sx5},${sy2} L${x4},${y2} L${sx6},${vc} L${x4},${y3} L${sx5},${sy3} L${x3},${y4} L${sx4},${sy4} L${hc},${h} L${sx3},${sy4} L${x2},${y4} L${sx2},${sy3} L${x1},${y3} L${sx1},${vc} z`;
            result += `<path d='${d}' fill='${fill}' stroke='${border.color}' stroke-width='${border.width}' stroke-dasharray='${border.strokeDasharray}' />`;
            break;
        }
        case "star12": {
            const adj = getAdjValue(node, "adj", 37500);
            const maxAdj = 50000 * SLIDE_FACTOR;
            const hd4 = h / 4;
            const wd4 = w / 4;
            const a = clamp(adj, 0, maxAdj);
            const dx1 = wd2 * Math.cos(0.5235987756);
            const dy1 = hd2 * Math.sin(1.0471975512);
            const x1 = hc - dx1;
            const x3 = w * 3 / 4;
            const x4 = hc + dx1;
            const y1 = vc - dy1;
            const y3 = h * 3 / 4;
            const y4 = vc + dy1;
            const iwd2 = wd2 * a / maxAdj;
            const ihd2 = hd2 * a / maxAdj;
            const sdx1 = iwd2 * Math.cos(0.2617993878);
            const sdx2 = iwd2 * Math.cos(0.7853981634);
            const sdx3 = iwd2 * Math.cos(1.308996939);
            const sdy1 = ihd2 * Math.sin(1.308996939);
            const sdy2 = ihd2 * Math.sin(0.7853981634);
            const sdy3 = ihd2 * Math.sin(0.2617993878);
            const sx1 = hc - sdx1;
            const sx2 = hc - sdx2;
            const sx3 = hc - sdx3;
            const sx4 = hc + sdx3;
            const sx5 = hc + sdx2;
            const sx6 = hc + sdx1;
            const sy1 = vc - sdy1;
            const sy2 = vc - sdy2;
            const sy3 = vc - sdy3;
            const sy4 = vc + sdy3;
            const sy5 = vc + sdy2;
            const sy6 = vc + sdy1;
            const d = `M0,${vc} L${sx1},${sy3} L${x1},${hd4} L${sx2},${sy2} L${wd4},${y1} L${sx3},${sy1} L${hc},0 L${sx4},${sy1} L${x3},${y1} L${sx5},${sy2} L${x4},${hd4} L${sx6},${sy3} L${w},${vc} L${sx6},${sy4} L${x4},${y3} L${sx5},${sy5} L${x3},${y4} L${sx4},${sy6} L${hc},${h} L${sx3},${sy6} L${wd4},${y4} L${sx2},${sy5} L${x1},${y3} L${sx1},${sy4} z`;
            result += `<path d='${d}' fill='${fill}' stroke='${border.color}' stroke-width='${border.width}' stroke-dasharray='${border.strokeDasharray}' />`;
            break;
        }
        case "star16": {
            const adj = getAdjValue(node, "adj", 37500);
            const maxAdj = 50000 * SLIDE_FACTOR;
            const a = clamp(adj, 0, maxAdj);
            const dx1 = wd2 * 92388 / 100000;
            const dx2 = wd2 * 70711 / 100000;
            const dx3 = wd2 * 38268 / 100000;
            const dy1 = hd2 * 92388 / 100000;
            const dy2 = hd2 * 70711 / 100000;
            const dy3 = hd2 * 38268 / 100000;
            const x1 = hc - dx1;
            const x2 = hc - dx2;
            const x3 = hc - dx3;
            const x4 = hc + dx3;
            const x5 = hc + dx2;
            const x6 = hc + dx1;
            const y1 = vc - dy1;
            const y2 = vc - dy2;
            const y3 = vc - dy3;
            const y4 = vc + dy3;
            const y5 = vc + dy2;
            const y6 = vc + dy1;
            const iwd2 = wd2 * a / maxAdj;
            const ihd2 = hd2 * a / maxAdj;
            const sdx1 = iwd2 * 98079 / 100000;
            const sdx2 = iwd2 * 83147 / 100000;
            const sdx3 = iwd2 * 55557 / 100000;
            const sdx4 = iwd2 * 19509 / 100000;
            const sdy1 = ihd2 * 98079 / 100000;
            const sdy2 = ihd2 * 83147 / 100000;
            const sdy3 = ihd2 * 55557 / 100000;
            const sdy4 = ihd2 * 19509 / 100000;
            const sx1 = hc - sdx1;
            const sx2 = hc - sdx2;
            const sx3 = hc - sdx3;
            const sx4 = hc - sdx4;
            const sx5 = hc + sdx4;
            const sx6 = hc + sdx3;
            const sx7 = hc + sdx2;
            const sx8 = hc + sdx1;
            const sy1 = vc - sdy1;
            const sy2 = vc - sdy2;
            const sy3 = vc - sdy3;
            const sy4 = vc - sdy4;
            const sy5 = vc + sdy4;
            const sy6 = vc + sdy3;
            const sy7 = vc + sdy2;
            const sy8 = vc + sdy1;
            const d = `M0,${vc} L${sx1},${sy4} L${x1},${y3} L${sx2},${sy3} L${x2},${y2} L${sx3},${sy2} L${x3},${y1} L${sx4},${sy1} L${hc},0 L${sx5},${sy1} L${x4},${y1} L${sx6},${sy2} L${x5},${y2} L${sx7},${sy3} L${x6},${y3} L${sx8},${sy4} L${w},${vc} L${sx8},${sy5} L${x6},${y4} L${sx7},${sy6} L${x5},${y5} L${sx6},${sy7} L${x4},${y6} L${sx5},${sy8} L${hc},${h} L${sx4},${sy8} L${x3},${y6} L${sx3},${sy7} L${x2},${y5} L${sx2},${sy6} L${x1},${y4} L${sx1},${sy5} z`;
            result += `<path d='${d}' fill='${fill}' stroke='${border.color}' stroke-width='${border.width}' stroke-dasharray='${border.strokeDasharray}' />`;
            break;
        }
        case "star24": {
            const adj = getAdjValue(node, "adj", 37500);
            const maxAdj = 50000 * SLIDE_FACTOR;
            const hd4 = h / 4;
            const wd4 = w / 4;
            const a = clamp(adj, 0, maxAdj);
            const dx1 = wd2 * Math.cos(0.2617993878);
            const dx2 = wd2 * Math.cos(0.5235987756);
            const dx3 = wd2 * Math.cos(0.7853981634);
            const dx4 = wd4;
            const dx5 = wd2 * Math.cos(1.308996939);
            const dy1 = hd2 * Math.sin(1.308996939);
            const dy2 = hd2 * Math.sin(1.0471975512);
            const dy3 = hd2 * Math.sin(0.7853981634);
            const dy4 = hd4;
            const dy5 = hd2 * Math.sin(0.2617993878);
            const x1 = hc - dx1;
            const x2 = hc - dx2;
            const x3 = hc - dx3;
            const x4 = hc - dx4;
            const x5 = hc - dx5;
            const x6 = hc + dx5;
            const x7 = hc + dx4;
            const x8 = hc + dx3;
            const x9 = hc + dx2;
            const x10 = hc + dx1;
            const y1 = vc - dy1;
            const y2 = vc - dy2;
            const y3 = vc - dy3;
            const y4 = vc - dy4;
            const y5 = vc - dy5;
            const y6 = vc + dy5;
            const y7 = vc + dy4;
            const y8 = vc + dy3;
            const y9 = vc + dy2;
            const y10 = vc + dy1;
            const iwd2 = wd2 * a / maxAdj;
            const ihd2 = hd2 * a / maxAdj;
            const sdx1 = iwd2 * 99144 / 100000;
            const sdx2 = iwd2 * 92388 / 100000;
            const sdx3 = iwd2 * 79335 / 100000;
            const sdx4 = iwd2 * 60876 / 100000;
            const sdx5 = iwd2 * 38268 / 100000;
            const sdx6 = iwd2 * 13053 / 100000;
            const sdy1 = ihd2 * 99144 / 100000;
            const sdy2 = ihd2 * 92388 / 100000;
            const sdy3 = ihd2 * 79335 / 100000;
            const sdy4 = ihd2 * 60876 / 100000;
            const sdy5 = ihd2 * 38268 / 100000;
            const sdy6 = ihd2 * 13053 / 100000;
            const sx1 = hc - sdx1;
            const sx2 = hc - sdx2;
            const sx3 = hc - sdx3;
            const sx4 = hc - sdx4;
            const sx5 = hc - sdx5;
            const sx6 = hc - sdx6;
            const sx7 = hc + sdx6;
            const sx8 = hc + sdx5;
            const sx9 = hc + sdx4;
            const sx10 = hc + sdx3;
            const sx11 = hc + sdx2;
            const sx12 = hc + sdx1;
            const sy1 = vc - sdy1;
            const sy2 = vc - sdy2;
            const sy3 = vc - sdy3;
            const sy4 = vc - sdy4;
            const sy5 = vc - sdy5;
            const sy6 = vc - sdy6;
            const sy7 = vc + sdy6;
            const sy8 = vc + sdy5;
            const sy9 = vc + sdy4;
            const sy10 = vc + sdy3;
            const sy11 = vc + sdy2;
            const sy12 = vc + sdy1;
            const d = `M0,${vc} L${sx1},${sy6} L${x1},${y5} L${sx2},${sy5} L${x2},${y4} L${sx3},${sy4} L${x3},${y3} L${sx4},${sy3} L${x4},${y2} L${sx5},${sy2} L${x5},${y1} L${sx6},${sy1} L${hc},${0} L${sx7},${sy1} L${x6},${y1} L${sx8},${sy2} L${x7},${y2} L${sx9},${sy3} L${x8},${y3} L${sx10},${sy4} L${x9},${y4} L${sx11},${sy5} L${x10},${y5} L${sx12},${sy6} L${w},${vc} L${sx12},${sy7} L${x10},${y6} L${sx11},${sy8} L${x9},${y7} L${sx10},${sy9} L${x8},${y8} L${sx9},${sy10} L${x7},${y9} L${sx8},${sy11} L${x6},${y10} L${sx7},${sy12} L${hc},${h} L${sx6},${sy12} L${x5},${y10} L${sx5},${sy11} L${x4},${y9} L${sx4},${sy10} L${x3},${y8} L${sx3},${sy9} L${x2},${y7} L${sx2},${sy8} L${x1},${y6} L${sx1},${sy7} z`;
            result += `<path d='${d}' fill='${fill}' stroke='${border.color}' stroke-width='${border.width}' stroke-dasharray='${border.strokeDasharray}' />`;
            break;
        }
        case "star32": {
            const adj = getAdjValue(node, "adj", 37500);
            const maxAdj = 50000 * SLIDE_FACTOR;
            const a = clamp(adj, 0, maxAdj);
            const dx1 = wd2 * 98079 / 100000;
            const dx2 = wd2 * 92388 / 100000;
            const dx3 = wd2 * 83147 / 100000;
            const dx4 = wd2 * Math.cos(0.7853981634);
            const dx5 = wd2 * 55557 / 100000;
            const dx6 = wd2 * 38268 / 100000;
            const dx7 = wd2 * 19509 / 100000;
            const dy1 = hd2 * 98079 / 100000;
            const dy2 = hd2 * 92388 / 100000;
            const dy3 = hd2 * 83147 / 100000;
            const dy4 = hd2 * Math.sin(0.7853981634);
            const dy5 = hd2 * 55557 / 100000;
            const dy6 = hd2 * 38268 / 100000;
            const dy7 = hd2 * 19509 / 100000;
            const x1 = hc - dx1;
            const x2 = hc - dx2;
            const x3 = hc - dx3;
            const x4 = hc - dx4;
            const x5 = hc - dx5;
            const x6 = hc - dx6;
            const x7 = hc - dx7;
            const x8 = hc + dx7;
            const x9 = hc + dx6;
            const x10 = hc + dx5;
            const x11 = hc + dx4;
            const x12 = hc + dx3;
            const x13 = hc + dx2;
            const x14 = hc + dx1;
            const y1 = vc - dy1;
            const y2 = vc - dy2;
            const y3 = vc - dy3;
            const y4 = vc - dy4;
            const y5 = vc - dy5;
            const y6 = vc - dy6;
            const y7 = vc - dy7;
            const y8 = vc + dy7;
            const y9 = vc + dy6;
            const y10 = vc + dy5;
            const y11 = vc + dy4;
            const y12 = vc + dy3;
            const y13 = vc + dy2;
            const y14 = vc + dy1;
            const iwd2 = wd2 * a / maxAdj;
            const ihd2 = hd2 * a / maxAdj;
            const sdx1 = iwd2 * 99518 / 100000;
            const sdx2 = iwd2 * 95694 / 100000;
            const sdx3 = iwd2 * 88192 / 100000;
            const sdx4 = iwd2 * 77301 / 100000;
            const sdx5 = iwd2 * 63439 / 100000;
            const sdx6 = iwd2 * 47140 / 100000;
            const sdx7 = iwd2 * 29028 / 100000;
            const sdx8 = iwd2 * 9802 / 100000;
            const sdy1 = ihd2 * 99518 / 100000;
            const sdy2 = ihd2 * 95694 / 100000;
            const sdy3 = ihd2 * 88192 / 100000;
            const sdy4 = ihd2 * 77301 / 100000;
            const sdy5 = ihd2 * 63439 / 100000;
            const sdy6 = ihd2 * 47140 / 100000;
            const sdy7 = ihd2 * 29028 / 100000;
            const sdy8 = ihd2 * 9802 / 100000;
            const sx1 = hc - sdx1;
            const sx2 = hc - sdx2;
            const sx3 = hc - sdx3;
            const sx4 = hc - sdx4;
            const sx5 = hc - sdx5;
            const sx6 = hc - sdx6;
            const sx7 = hc - sdx7;
            const sx8 = hc - sdx8;
            const sx9 = hc + sdx8;
            const sx10 = hc + sdx7;
            const sx11 = hc + sdx6;
            const sx12 = hc + sdx5;
            const sx13 = hc + sdx4;
            const sx14 = hc + sdx3;
            const sx15 = hc + sdx2;
            const sx16 = hc + sdx1;
            const sy1 = vc - sdy1;
            const sy2 = vc - sdy2;
            const sy3 = vc - sdy3;
            const sy4 = vc - sdy4;
            const sy5 = vc - sdy5;
            const sy6 = vc - sdy6;
            const sy7 = vc - sdy7;
            const sy8 = vc - sdy8;
            const sy9 = vc + sdy8;
            const sy10 = vc + sdy7;
            const sy11 = vc + sdy6;
            const sy12 = vc + sdy5;
            const sy13 = vc + sdy4;
            const sy14 = vc + sdy3;
            const sy15 = vc + sdy2;
            const sy16 = vc + sdy1;
            const d = `M0,${vc} L${sx1},${sy8} L${x1},${y7} L${sx2},${sy7} L${x2},${y6} L${sx3},${sy6} L${x3},${y5} L${sx4},${sy5} L${x4},${y4} L${sx5},${sy4} L${x5},${y3} L${sx6},${sy3} L${x6},${y2} L${sx7},${sy2} L${x7},${y1} L${sx8},${sy1} L${hc},${0} L${sx9},${sy1} L${x8},${y1} L${sx10},${sy2} L${x9},${y2} L${sx11},${sy3} L${x10},${y3} L${sx12},${sy4} L${x11},${y4} L${sx13},${sy5} L${x12},${y5} L${sx14},${sy6} L${x13},${y6} L${sx15},${sy7} L${x14},${y7} L${sx16},${sy8} L${w},${vc} L${sx16},${sy9} L${x14},${y8} L${sx15},${sy10} L${x13},${y9} L${sx14},${sy11} L${x12},${y10} L${sx13},${sy12} L${x11},${y11} L${sx12},${sy13} L${x10},${y12} L${sx11},${sy14} L${x9},${y13} L${sx10},${sy15} L${x8},${y14} L${sx9},${sy16} L${hc},${h} L${sx8},${sy16} L${x7},${y14} L${sx7},${sy15} L${x6},${y13} L${sx6},${sy14} L${x5},${y12} L${sx5},${sy13} L${x4},${y11} L${sx4},${sy12} L${x3},${y10} L${sx3},${sy11} L${x2},${y9} L${sx2},${sy10} L${x1},${y8} L${sx1},${sy9} z`;
            result += `<path d='${d}' fill='${fill}' stroke='${border.color}' stroke-width='${border.width}' stroke-dasharray='${border.strokeDasharray}' />`;
            break;
        }
    }
    return result;
}
function getAdjValue(node, name, defaultValue) {
    const shapAdjst = PPTXXmlUtils.getTextByPathList(node, ["p:spPr", "a:prstGeom", "a:avLst", "a:gd"]);
    if (shapAdjst !== undefined) {
        if (Array.isArray(shapAdjst)) {
            for (let key of Object.keys(shapAdjst)) {
                if (shapAdjst[key] && shapAdjst[key]["attrs"] && shapAdjst[key]["attrs"]["name"] === name) {
                    return parseInt(shapAdjst[key]["attrs"]["fmla"].substr(4)) * SLIDE_FACTOR;
                }
            }
        }
        else if (shapAdjst["attrs"] && shapAdjst["attrs"]["name"] === name) {
            return parseInt(shapAdjst["attrs"]["fmla"].substr(4)) * SLIDE_FACTOR;
        }
    }
    return defaultValue * SLIDE_FACTOR;
}
function clamp(value, min, max) {
    return value < min ? min : value > max ? max : value;
}

function renderMathSymbol(shapType, w, h, imgFillFlg, grndFillFlg, fillColor, border, shpId, node) {
    let result = "";
    const shapAdjst_ary = PPTXXmlUtils.getTextByPathList(node, ["p:spPr", "a:prstGeom", "a:avLst", "a:gd"]);
    let sAdj1, adj1;
    let sAdj2, adj2;
    let sAdj3, adj3;
    if (shapAdjst_ary !== undefined) {
        if (shapAdjst_ary.constructor === Array) {
            for (const item of shapAdjst_ary) {
                const sAdj_name = PPTXXmlUtils.getTextByPathList(item, ["attrs", "name"]);
                if (sAdj_name == "adj1") {
                    sAdj1 = PPTXXmlUtils.getTextByPathList(item, ["attrs", "fmla"]);
                    adj1 = parseInt(sAdj1.substr(4));
                }
                else if (sAdj_name == "adj2") {
                    sAdj2 = PPTXXmlUtils.getTextByPathList(item, ["attrs", "fmla"]);
                    adj2 = parseInt(sAdj2.substr(4));
                }
                else if (sAdj_name == "adj3") {
                    sAdj3 = PPTXXmlUtils.getTextByPathList(item, ["attrs", "fmla"]);
                    adj3 = parseInt(sAdj3.substr(4));
                }
            }
        }
        else {
            sAdj1 = PPTXXmlUtils.getTextByPathList(shapAdjst_ary, ["attrs", "fmla"]);
            adj1 = parseInt(sAdj1.substr(4));
        }
    }
    const cnstVal1 = 50000 * SLIDE_FACTOR$1;
    const cnstVal2 = 100000 * SLIDE_FACTOR$1;
    const cnstVal3 = 200000 * SLIDE_FACTOR$1;
    let dVal;
    const hc = w / 2, vc = h / 2, hd2 = h / 2;
    if (shapType == "mathNotEqual") {
        if (shapAdjst_ary === undefined) {
            adj1 = 23520 * SLIDE_FACTOR$1;
            adj2 = 110 * Math.PI / 180;
            adj3 = 11760 * SLIDE_FACTOR$1;
        }
        else {
            adj1 = adj1 * SLIDE_FACTOR$1;
            adj2 = (adj2 / 60000) * Math.PI / 180;
            adj3 = adj3 * SLIDE_FACTOR$1;
        }
        let a1, crAng, a2a1, maxAdj3, a3, dy1, dy2, dx1, x1, x8, y2, y3, y1, y4, cadj2, xadj2, len, bhw, bhw2, x7, dx67, x6, dx57, x5, dx47, x4, dx37, x3, rx7, rx6, rx5, rx4, rx3, dx7, rxt, lxt, rx, lx, dy3, dy4, ry, ly, dlx, drx, dly, dry;
        const angVal1 = 70 * Math.PI / 180, angVal2 = 110 * Math.PI / 180;
        const cnstVal4 = 73490 * SLIDE_FACTOR$1;
        a1 = (adj1 < 0) ? 0 : (adj1 > cnstVal1) ? cnstVal1 : adj1;
        crAng = (adj2 < angVal1) ? angVal1 : (adj2 > angVal2) ? angVal2 : adj2;
        a2a1 = a1 * 2;
        maxAdj3 = cnstVal2 - a2a1;
        a3 = (adj3 < 0) ? 0 : (adj3 > maxAdj3) ? maxAdj3 : adj3;
        dy1 = h * a1 / cnstVal2;
        dy2 = h * a3 / cnstVal3;
        dx1 = w * cnstVal4 / cnstVal3;
        x1 = hc - dx1;
        x8 = hc + dx1;
        y2 = vc - dy2;
        y3 = vc + dy2;
        y1 = y2 - dy1;
        y4 = y3 + dy1;
        cadj2 = crAng - Math.PI / 2;
        xadj2 = hd2 * Math.tan(cadj2);
        len = Math.sqrt(xadj2 * xadj2 + hd2 * hd2);
        bhw = len * dy1 / hd2;
        bhw2 = bhw / 2;
        x7 = hc + xadj2 - bhw2;
        dx67 = xadj2 * y1 / hd2;
        x6 = x7 - dx67;
        dx57 = xadj2 * y2 / hd2;
        x5 = x7 - dx57;
        dx47 = xadj2 * y3 / hd2;
        x4 = x7 - dx47;
        dx37 = xadj2 * y4 / hd2;
        x3 = x7 - dx37;
        rx7 = x7 + bhw;
        rx6 = x6 + bhw;
        rx5 = x5 + bhw;
        rx4 = x4 + bhw;
        rx3 = x3 + bhw;
        dx7 = dy1 * hd2 / len;
        rxt = x7 + dx7;
        lxt = rx7 - dx7;
        rx = (cadj2 > 0) ? rxt : rx7;
        lx = (cadj2 > 0) ? x7 : lxt;
        dy3 = dy1 * xadj2 / len;
        dy4 = -dy3;
        ry = (cadj2 > 0) ? dy3 : 0;
        ly = (cadj2 > 0) ? 0 : dy4;
        dlx = w - rx;
        drx = w - lx;
        dly = h - ry;
        dry = h - ly;
        dVal = `M${x1},${y1} L${x6},${y1} L${lx},${ly} L${rx},${ry} L${rx6},${y1} L${x8},${y1} L${x8},${y2} L${rx5},${y2} L${rx4},${y3} L${x8},${y3} L${x8},${y4} L${rx3},${y4} L${drx},${dry} L${dlx},${dly} L${x3},${y4} L${x1},${y4} L${x1},${y3} L${x4},${y3} L${x5},${y2} L${x1},${y2} z`;
    }
    else if (shapType == "mathDivide") {
        if (shapAdjst_ary === undefined) {
            adj1 = 23520 * SLIDE_FACTOR$1;
            adj2 = 5880 * SLIDE_FACTOR$1;
            adj3 = 11760 * SLIDE_FACTOR$1;
        }
        else {
            adj1 = adj1 * SLIDE_FACTOR$1;
            adj2 = adj2 * SLIDE_FACTOR$1;
            adj3 = adj3 * SLIDE_FACTOR$1;
        }
        let a1, ma1, ma3h, ma3w, maxAdj3, a3, m4a3, maxAdj2, a2, dy1, yg, rad, dx1, y3, y4, a, y2, y1, y5, x1, x3;
        const cnstVal4 = 1000 * SLIDE_FACTOR$1;
        const cnstVal5 = 36745 * SLIDE_FACTOR$1;
        const cnstVal6 = 73490 * SLIDE_FACTOR$1;
        a1 = (adj1 < cnstVal4) ? cnstVal4 : (adj1 > cnstVal5) ? cnstVal5 : adj1;
        ma1 = -a1;
        ma3h = (cnstVal6 + ma1) / 4;
        ma3w = cnstVal5 * w / h;
        maxAdj3 = (ma3h < ma3w) ? ma3h : ma3w;
        a3 = (adj3 < cnstVal4) ? cnstVal4 : (adj3 > maxAdj3) ? maxAdj3 : adj3;
        m4a3 = -4 * a3;
        maxAdj2 = cnstVal6 + m4a3 - a1;
        a2 = (adj2 < 0) ? 0 : (adj2 > maxAdj2) ? maxAdj2 : adj2;
        dy1 = h * a1 / cnstVal3;
        yg = h * a2 / cnstVal2;
        rad = h * a3 / cnstVal2;
        dx1 = w * cnstVal6 / cnstVal3;
        y3 = vc - dy1;
        y4 = vc + dy1;
        a = yg + rad;
        y2 = y3 - a;
        y1 = y2 - rad;
        y5 = h - y1;
        x1 = hc - dx1;
        x3 = hc + dx1;
        const cd4 = 90, c3d4 = 270;
        const cX1 = hc - Math.cos(c3d4 * Math.PI / 180) * rad;
        const cY1 = y1 - Math.sin(c3d4 * Math.PI / 180) * rad;
        const cX2 = hc - Math.cos(Math.PI / 2) * rad;
        const cY2 = y5 - Math.sin(Math.PI / 2) * rad;
        dVal = `M${hc},${y1}${shapeArc(cX1, cY1, rad, rad, c3d4, c3d4 + 360, false).replace("M", "L")} z M${hc},${y5}${shapeArc(cX2, cY2, rad, rad, cd4, cd4 + 360, false).replace("M", "L")} z M${x1},${y3} L${x3},${y3} L${x3},${y4} L${x1},${y4} z`;
    }
    else if (shapType == "mathEqual") {
        if (shapAdjst_ary === undefined) {
            adj1 = 23520 * SLIDE_FACTOR$1;
            adj2 = 11760 * SLIDE_FACTOR$1;
        }
        else {
            adj1 = adj1 * SLIDE_FACTOR$1;
            adj2 = adj2 * SLIDE_FACTOR$1;
        }
        const cnstVal5 = 36745 * SLIDE_FACTOR$1;
        const cnstVal6 = 73490 * SLIDE_FACTOR$1;
        let a1, a2a1, mAdj2, a2, dy1, dy2, dx1, y2, y3, y1, y4, x1, x2;
        a1 = (adj1 < 0) ? 0 : (adj1 > cnstVal5) ? cnstVal5 : adj1;
        a2a1 = a1 * 2;
        mAdj2 = cnstVal2 - a2a1;
        a2 = (adj2 < 0) ? 0 : (adj2 > mAdj2) ? mAdj2 : adj2;
        dy1 = h * a1 / cnstVal2;
        dy2 = h * a2 / cnstVal3;
        dx1 = w * cnstVal6 / cnstVal3;
        y2 = vc - dy2;
        y3 = vc + dy2;
        y1 = y2 - dy1;
        y4 = y3 + dy1;
        x1 = hc - dx1;
        x2 = hc + dx1;
        dVal = `M${x1},${y1} L${x2},${y1} L${x2},${y2} L${x1},${y2} zM${x1},${y3} L${x2},${y3} L${x2},${y4} L${x1},${y4} z`;
    }
    else if (shapType == "mathMinus") {
        if (shapAdjst_ary === undefined) {
            adj1 = 23520 * SLIDE_FACTOR$1;
        }
        else {
            adj1 = adj1 * SLIDE_FACTOR$1;
        }
        const cnstVal6 = 73490 * SLIDE_FACTOR$1;
        let a1, dy1, dx1, y1, y2, x1, x2;
        a1 = (adj1 < 0) ? 0 : (adj1 > cnstVal2) ? cnstVal2 : adj1;
        dy1 = h * a1 / cnstVal3;
        dx1 = w * cnstVal6 / cnstVal3;
        y1 = vc - dy1;
        y2 = vc + dy1;
        x1 = hc - dx1;
        x2 = hc + dx1;
        dVal = `M${x1},${y1} L${x2},${y1} L${x2},${y2} L${x1},${y2} z`;
    }
    else if (shapType == "mathMultiply") {
        if (shapAdjst_ary === undefined) {
            adj1 = 23520 * SLIDE_FACTOR$1;
        }
        else {
            adj1 = adj1 * SLIDE_FACTOR$1;
        }
        const cnstVal6 = 51965 * SLIDE_FACTOR$1;
        let a1, th, a, sa, ca, ta, dl, rw, lM, xM, yM, dxAM, dyAM, xA, yA, xB, yB, xBC, yBC, yC, xD, xE, yFE, xFE, xF, xL, yG, yH, yI;
        const ss = Math.min(w, h);
        a1 = (adj1 < 0) ? 0 : (adj1 > cnstVal6) ? cnstVal6 : adj1;
        th = ss * a1 / cnstVal2;
        a = Math.atan(h / w);
        sa = 1 * Math.sin(a);
        ca = 1 * Math.cos(a);
        ta = 1 * Math.tan(a);
        dl = Math.sqrt(w * w + h * h);
        rw = dl * cnstVal6 / cnstVal2;
        lM = dl - rw;
        xM = ca * lM / 2;
        yM = sa * lM / 2;
        dxAM = sa * th / 2;
        dyAM = ca * th / 2;
        xA = xM - dxAM;
        yA = yM + dyAM;
        xB = xM + dxAM;
        yB = yM - dyAM;
        xBC = hc - xB;
        yBC = xBC * ta;
        yC = yBC + yB;
        xD = w - xB;
        xE = w - xA;
        yFE = vc - yA;
        xFE = yFE / ta;
        xF = xE - xFE;
        xL = xA + xFE;
        yG = h - yA;
        yH = h - yB;
        yI = h - yC;
        dVal = `M${xA},${yA} L${xB},${yB} L${hc},${yC} L${xD},${yB} L${xE},${yA} L${xF},${vc} L${xE},${yG} L${xD},${yH} L${hc},${yI} L${xB},${yH} L${xA},${yG} L${xL},${vc} z`;
    }
    else if (shapType == "mathPlus") {
        if (shapAdjst_ary === undefined) {
            adj1 = 23520 * SLIDE_FACTOR$1;
        }
        else {
            adj1 = adj1 * SLIDE_FACTOR$1;
        }
        const cnstVal6 = 73490 * SLIDE_FACTOR$1;
        const ss = Math.min(w, h);
        let a1, dx1, dy1, dx2, x1, x2, x3, x4, y1, y2, y3, y4;
        a1 = (adj1 < 0) ? 0 : (adj1 > cnstVal6) ? cnstVal6 : adj1;
        dx1 = w * cnstVal6 / cnstVal3;
        dy1 = h * cnstVal6 / cnstVal3;
        dx2 = ss * a1 / cnstVal3;
        x1 = hc - dx1;
        x2 = hc - dx2;
        x3 = hc + dx2;
        x4 = hc + dx1;
        y1 = vc - dy1;
        y2 = vc - dx2;
        y3 = vc + dx2;
        y4 = vc + dy1;
        dVal = `M${x1},${y2} L${x2},${y2} L${x2},${y1} L${x3},${y1} L${x3},${y2} L${x4},${y2} L${x4},${y3} L${x3},${y3} L${x3},${y4} L${x2},${y4} L${x2},${y3} L${x1},${y3} z`;
    }
    result += `<path d='${dVal}' fill='${(!imgFillFlg ? (grndFillFlg ? "url(#linGrd_" + shpId + ")" : fillColor) : "url(#imgPtrn_" + shpId + ")")}' stroke='${border.color}' stroke-width='${border.width}' stroke-dasharray='${border.strokeDasharray}' />`;
    return result;
}

function renderBracket(shapType, w, h, imgFillFlg, grndFillFlg, fillColor, border, shpId, node) {
    let result = "";
    let dVal = "";
    if (shapType === "bracePair") {
        const shapAdjst = PPTXXmlUtils.getTextByPathList(node, ["p:spPr", "a:prstGeom", "a:avLst", "a:gd", "attrs", "fmla"]);
        let adj = 8333 * SLIDE_FACTOR$1;
        const cnstVal1 = 25000 * SLIDE_FACTOR$1;
        const cnstVal2 = 50000 * SLIDE_FACTOR$1;
        const cnstVal3 = 100000 * SLIDE_FACTOR$1;
        if (shapAdjst !== undefined) {
            adj = parseInt(shapAdjst.substr(4)) * SLIDE_FACTOR$1;
        }
        let vc = h / 2, cd = 360, cd2 = 180, cd4 = 90, c3d4 = 270, a, x1, x2, x3, x4, y2, y3, y4;
        if (adj < 0)
            a = 0;
        else if (adj > cnstVal1)
            a = cnstVal1;
        else
            a = adj;
        const minWH = Math.min(w, h);
        x1 = minWH * a / cnstVal3;
        x2 = minWH * a / cnstVal2;
        x3 = w - x2;
        x4 = w - x1;
        y2 = vc - x1;
        y3 = vc + x1;
        y4 = h - x1;
        dVal = `M${x2},${h}${shapeArc(x2, y4, x1, x1, cd4, cd2, false).replace("M", "L")} L${x1},${y3}${shapeArc(0, y3, x1, x1, 0, (-cd4), false).replace("M", "L")}${shapeArc(0, y2, x1, x1, cd4, 0, false).replace("M", "L")} L${x1},${x1}${shapeArc(x2, x1, x1, x1, cd2, c3d4, false).replace("M", "L")} M${x3},${0}${shapeArc(x3, x1, x1, x1, c3d4, cd, false).replace("M", "L")} L${x4},${y2}${shapeArc(w, y2, x1, x1, cd2, cd4, false).replace("M", "L")}${shapeArc(w, y3, x1, x1, c3d4, cd2, false).replace("M", "L")} L${x4},${y4}${shapeArc(x3, y4, x1, x1, 0, cd4, false).replace("M", "L")}`;
    }
    else if (shapType === "leftBrace") {
        const shapAdjst_ary = PPTXXmlUtils.getTextByPathList(node, ["p:spPr", "a:prstGeom", "a:avLst", "a:gd"]);
        let sAdj1, adj1 = 8333 * SLIDE_FACTOR$1;
        let sAdj2, adj2 = 50000 * SLIDE_FACTOR$1;
        const cnstVal2 = 100000 * SLIDE_FACTOR$1;
        if (shapAdjst_ary !== undefined) {
            for (const i of shapAdjst_ary.keys()) {
                const sAdj_name = PPTXXmlUtils.getTextByPathList(shapAdjst_ary[i], ["attrs", "name"]);
                if (sAdj_name == "adj1") {
                    sAdj1 = PPTXXmlUtils.getTextByPathList(shapAdjst_ary[i], ["attrs", "fmla"]);
                    adj1 = parseInt(sAdj1.substr(4)) * SLIDE_FACTOR$1;
                }
                else if (sAdj_name == "adj2") {
                    sAdj2 = PPTXXmlUtils.getTextByPathList(shapAdjst_ary[i], ["attrs", "fmla"]);
                    adj2 = parseInt(sAdj2.substr(4)) * SLIDE_FACTOR$1;
                }
            }
        }
        let cd2 = 180, cd4 = 90, c3d4 = 270, a1, a2, q1, q2, q3, y1, y2, y3, y4;
        if (adj2 < 0)
            a2 = 0;
        else if (adj2 > cnstVal2)
            a2 = cnstVal2;
        else
            a2 = adj2;
        const minWH = Math.min(w, h);
        q1 = cnstVal2 - a2;
        if (q1 < a2)
            q2 = q1;
        else
            q2 = a2;
        q3 = q2 / 2;
        const maxAdj1 = q3 * h / minWH;
        if (adj1 < 0)
            a1 = 0;
        else if (adj1 > maxAdj1)
            a1 = maxAdj1;
        else
            a1 = adj1;
        y1 = minWH * a1 / cnstVal2;
        y3 = h * a2 / cnstVal2;
        y2 = y3 - y1;
        y4 = y3 + y1;
        dVal = `M${w},${h}${shapeArc(w, h - y1, w / 2, y1, cd4, cd2, false).replace("M", "L")} L${w / 2},${y4}${shapeArc(0, y4, w / 2, y1, 0, (-cd4), false).replace("M", "L")}${shapeArc(0, y2, w / 2, y1, cd4, 0, false).replace("M", "L")} L${w / 2},${y1}${shapeArc(w, y1, w / 2, y1, cd2, c3d4, false).replace("M", "L")}`;
    }
    else if (shapType === "rightBrace") {
        const shapAdjst_ary = PPTXXmlUtils.getTextByPathList(node, ["p:spPr", "a:prstGeom", "a:avLst", "a:gd"]);
        let sAdj1, adj1 = 8333 * SLIDE_FACTOR$1;
        let sAdj2, adj2 = 50000 * SLIDE_FACTOR$1;
        const cnstVal2 = 100000 * SLIDE_FACTOR$1;
        if (shapAdjst_ary !== undefined) {
            for (const i of shapAdjst_ary.keys()) {
                const sAdj_name = PPTXXmlUtils.getTextByPathList(shapAdjst_ary[i], ["attrs", "name"]);
                if (sAdj_name == "adj1") {
                    sAdj1 = PPTXXmlUtils.getTextByPathList(shapAdjst_ary[i], ["attrs", "fmla"]);
                    adj1 = parseInt(sAdj1.substr(4)) * SLIDE_FACTOR$1;
                }
                else if (sAdj_name == "adj2") {
                    sAdj2 = PPTXXmlUtils.getTextByPathList(shapAdjst_ary[i], ["attrs", "fmla"]);
                    adj2 = parseInt(sAdj2.substr(4)) * SLIDE_FACTOR$1;
                }
            }
        }
        let cd = 360, cd2 = 180, cd4 = 90, c3d4 = 270, a1, a2, q1, q2, q3, y1, y2, y3, y4;
        if (adj2 < 0)
            a2 = 0;
        else if (adj2 > cnstVal2)
            a2 = cnstVal2;
        else
            a2 = adj2;
        const minWH = Math.min(w, h);
        q1 = cnstVal2 - a2;
        if (q1 < a2)
            q2 = q1;
        else
            q2 = a2;
        q3 = q2 / 2;
        const maxAdj1 = q3 * h / minWH;
        if (adj1 < 0)
            a1 = 0;
        else if (adj1 > maxAdj1)
            a1 = maxAdj1;
        else
            a1 = adj1;
        y1 = minWH * a1 / cnstVal2;
        y3 = h * a2 / cnstVal2;
        y2 = y3 - y1;
        y4 = h - y1;
        dVal = `M${0},${0}${shapeArc(0, y1, w / 2, y1, c3d4, cd, false).replace("M", "L")} L${w / 2},${y2}${shapeArc(w, y2, w / 2, y1, cd2, cd4, false).replace("M", "L")}${shapeArc(w, y3 + y1, w / 2, y1, c3d4, cd2, false).replace("M", "L")} L${w / 2},${y4}${shapeArc(0, y4, w / 2, y1, 0, cd4, false).replace("M", "L")}`;
    }
    else if (shapType === "bracketPair") {
        const shapAdjst = PPTXXmlUtils.getTextByPathList(node, ["p:spPr", "a:prstGeom", "a:avLst", "a:gd", "attrs", "fmla"]);
        let adj = 16667 * SLIDE_FACTOR$1;
        const cnstVal1 = 50000 * SLIDE_FACTOR$1;
        const cnstVal2 = 100000 * SLIDE_FACTOR$1;
        if (shapAdjst !== undefined) {
            adj = parseInt(shapAdjst.substr(4)) * SLIDE_FACTOR$1;
        }
        let r = w, b = h, cd2 = 180, cd4 = 90, c3d4 = 270, a, x1, x2, y2;
        if (adj < 0)
            a = 0;
        else if (adj > cnstVal1)
            a = cnstVal1;
        else
            a = adj;
        x1 = Math.min(w, h) * a / cnstVal2;
        x2 = r - x1;
        y2 = b - x1;
        dVal = shapeArc(x1, x1, x1, x1, c3d4, cd2, false) +
            shapeArc(x1, y2, x1, x1, cd2, cd4, false).replace("M", "L") +
            shapeArc(x2, x1, x1, x1, c3d4, (c3d4 + cd4), false) +
            shapeArc(x2, y2, x1, x1, 0, cd4, false).replace("M", "L");
    }
    else if (shapType === "leftBracket") {
        const shapAdjst = PPTXXmlUtils.getTextByPathList(node, ["p:spPr", "a:prstGeom", "a:avLst", "a:gd", "attrs", "fmla"]);
        let adj = 8333 * SLIDE_FACTOR$1;
        const cnstVal1 = 50000 * SLIDE_FACTOR$1;
        const cnstVal2 = 100000 * SLIDE_FACTOR$1;
        const maxAdj = cnstVal1 * h / Math.min(w, h);
        if (shapAdjst !== undefined) {
            adj = parseInt(shapAdjst.substr(4)) * SLIDE_FACTOR$1;
        }
        let r = w, b = h, cd2 = 180, cd4 = 90, c3d4 = 270, a, y1, y2;
        if (adj < 0)
            a = 0;
        else if (adj > maxAdj)
            a = maxAdj;
        else
            a = adj;
        y1 = Math.min(w, h) * a / cnstVal2;
        if (y1 > w)
            y1 = w;
        y2 = b - y1;
        dVal = `M${r},${b}${shapeArc(y1, y2, y1, y1, cd4, cd2, false).replace("M", "L")} L${0},${y1}${shapeArc(y1, y1, y1, y1, cd2, c3d4, false).replace("M", "L")} L${r},${0}`;
    }
    else if (shapType === "rightBracket") {
        const shapAdjst = PPTXXmlUtils.getTextByPathList(node, ["p:spPr", "a:prstGeom", "a:avLst", "a:gd", "attrs", "fmla"]);
        let adj = 8333 * SLIDE_FACTOR$1;
        const cnstVal1 = 50000 * SLIDE_FACTOR$1;
        const cnstVal2 = 100000 * SLIDE_FACTOR$1;
        const maxAdj = cnstVal1 * h / Math.min(w, h);
        if (shapAdjst !== undefined) {
            adj = parseInt(shapAdjst.substr(4)) * SLIDE_FACTOR$1;
        }
        let cd = 360, cd4 = 90, c3d4 = 270, a, y1, y2, y3;
        if (adj < 0)
            a = 0;
        else if (adj > maxAdj)
            a = maxAdj;
        else
            a = adj;
        y1 = Math.min(w, h) * a / cnstVal2;
        y2 = h - y1;
        y3 = w - y1;
        dVal = `M${0},${h}${shapeArc(y3, y2, y1, y1, cd4, 0, false).replace("M", "L")} L${w},${h / 2}${shapeArc(y3, y1, y1, y1, cd, c3d4, false).replace("M", "L")} L${0},${0}`;
    }
    result += `<path d='${dVal}' fill='${(!imgFillFlg ? (grndFillFlg ? "url(#linGrd_" + shpId + ")" : fillColor) : "url(#imgPtrn_" + shpId + ")")}' stroke='${border.color}' stroke-width='${border.width}' stroke-dasharray='${border.strokeDasharray}' />`;
    return result;
}

function renderMiscShape(shapType, w, h, imgFillFlg, grndFillFlg, fillColor, border, shpId, node) {
    let result = "";
    let dVal = "";
    if (shapType === "smileyFace") {
        const shapAdjst = PPTXXmlUtils.getTextByPathList(node, ["p:spPr", "a:prstGeom", "a:avLst", "a:gd", "attrs", "fmla"]);
        const refr = SLIDE_FACTOR$1;
        let adj = 4653 * refr;
        if (shapAdjst !== undefined) {
            adj = parseInt(shapAdjst.substr(4)) * refr;
        }
        const cnstVal1 = 50000 * refr;
        const cnstVal2 = 100000 * refr;
        const cnstVal3 = 4653 * refr;
        let a, x1, x2, x3, x4, y1, y3, dy2, y2, y4, dy3, y5, wR, hR, wd2, hd2;
        wd2 = w / 2;
        hd2 = h / 2;
        a = (adj < -cnstVal3) ? -cnstVal3 : (adj > cnstVal3) ? cnstVal3 : adj;
        x1 = w * 4969 / 21699;
        x2 = w * 6215 / 21600;
        x3 = w * 13135 / 21600;
        x4 = w * 16640 / 21600;
        y1 = h * 7570 / 21600;
        y3 = h * 16515 / 21600;
        dy2 = h * a / cnstVal2;
        y2 = y3 - dy2;
        y4 = y3 + dy2;
        dy3 = h * a / cnstVal1;
        y5 = y4 + dy3;
        wR = w * 1125 / 21600;
        hR = h * 1125 / 21600;
        const cX1 = x2 - wR * Math.cos(Math.PI);
        const cY1 = y1 - hR * Math.sin(Math.PI);
        const cX2 = x3 - wR * Math.cos(Math.PI);
        dVal =
            `${shapeArc(cX1, cY1, wR, hR, 180, 540, false)}${shapeArc(cX2, cY1, wR, hR, 180, 540, false)} M${x1},${y2} Q${wd2},${y5} ${x4},${y2} Q${wd2},${y5} ${x1},${y2} M${0},${hd2}${shapeArc(wd2, hd2, wd2, hd2, 180, 540, false).replace("M", "L")} z`;
    }
    else if (shapType === "verticalScroll" || shapType === "horizontalScroll") {
        const shapAdjst = PPTXXmlUtils.getTextByPathList(node, ["p:spPr", "a:prstGeom", "a:avLst", "a:gd", "attrs", "fmla"]);
        const refr = SLIDE_FACTOR$1;
        let adj = 12500 * refr;
        if (shapAdjst !== undefined) {
            adj = parseInt(shapAdjst.substr(4)) * refr;
        }
        const cnstVal1 = 25000 * refr;
        const cnstVal2 = 100000 * refr;
        const ss = Math.min(w, h);
        const t = 0, l = 0, b = h, r = w;
        let a, ch, ch2, ch4;
        a = (adj < 0) ? 0 : (adj > cnstVal1) ? cnstVal1 : adj;
        ch = ss * a / cnstVal2;
        ch2 = ch / 2;
        ch4 = ch / 4;
        if (shapType === "verticalScroll") {
            let x3, x4, x6, x7, x5, y3, y4;
            x3 = ch + ch2;
            x4 = ch + ch;
            x6 = r - ch;
            x7 = r - ch2;
            x5 = x6 - ch2;
            y3 = b - ch;
            y4 = b - ch2;
            dVal = `M${ch},${y3} L${ch},${ch2}${shapeArc(x3, ch2, ch2, ch2, 180, 270, false).replace("M", "L")} L${x7},${t}${shapeArc(x7, ch2, ch2, ch2, 270, 450, false).replace("M", "L")} L${x6},${ch} L${x6},${y4}${shapeArc(x5, y4, ch2, ch2, 0, 90, false).replace("M", "L")} L${ch2},${b}${shapeArc(ch2, y4, ch2, ch2, 90, 270, false).replace("M", "L")} z M${x3},${t}${shapeArc(x3, ch2, ch2, ch2, 270, 450, false).replace("M", "L")}${shapeArc(x3, x3 / 2, ch4, ch4, 90, 270, false).replace("M", "L")} L${x4},${ch2} M${x6},${ch} L${x3},${ch} M${ch},${y4}${shapeArc(ch2, y4, ch2, ch2, 0, 270, false).replace("M", "L")}${shapeArc(ch2, (y4 + y3) / 2, ch4, ch4, 270, 450, false).replace("M", "L")} z M${ch},${y4} L${ch},${y3}`;
        }
        else if (shapType === "horizontalScroll") {
            let y3, y4, y6, y7, y5, x3, x4;
            y3 = ch + ch2;
            y4 = ch + ch;
            y6 = b - ch;
            y7 = b - ch2;
            y5 = y6 - ch2;
            x3 = r - ch;
            x4 = r - ch2;
            dVal = `M${l},${y3}${shapeArc(ch2, y3, ch2, ch2, 180, 270, false).replace("M", "L")} L${x3},${ch} L${x3},${ch2}${shapeArc(x4, ch2, ch2, ch2, 180, 360, false).replace("M", "L")} L${r},${y5}${shapeArc(x4, y5, ch2, ch2, 0, 90, false).replace("M", "L")} L${ch},${y6} L${ch},${y7}${shapeArc(ch2, y7, ch2, ch2, 0, 180, false).replace("M", "L")} zM${x4},${ch}${shapeArc(x4, ch2, ch2, ch2, 90, -180, false).replace("M", "L")}${shapeArc((x3 + x4) / 2, ch2, ch4, ch4, 180, 0, false).replace("M", "L")} z M${x4},${ch} L${x3},${ch} M${ch2},${y4} L${ch2},${y3}${shapeArc(y3 / 2, y3, ch4, ch4, 180, 360, false).replace("M", "L")}${shapeArc(ch2, y3, ch2, ch2, 0, 180, false).replace("M", "L")} M${ch},${y3} L${ch},${y6}`;
        }
    }
    result += `<path d='${dVal}' fill='${(!imgFillFlg ? (grndFillFlg ? "url(#linGrd_" + shpId + ")" : fillColor) : "url(#imgPtrn_" + shpId + ")")}' stroke='${border.color}' stroke-width='${border.width}' stroke-dasharray='${border.strokeDasharray}' />`;
    return result;
}

function renderPieShape(shapType, w, h, imgFillFlg, grndFillFlg, fillColor, border, shpId, node, oShadowSvgUrlStr) {
    let result = "";
    let dVal = "";
    if (shapType === "pie" || shapType === "pieWedge" || shapType === "arc") {
        const shapAdjst = PPTXXmlUtils.getTextByPathList(node, ["p:spPr", "a:prstGeom", "a:avLst", "a:gd"]);
        let adj1, adj2, H, shapAdjst1, shapAdjst2, isClose;
        if (shapType === "pie") {
            adj1 = 0;
            adj2 = 270;
            H = h;
            isClose = true;
        }
        else if (shapType === "pieWedge") {
            adj1 = 180;
            adj2 = 270;
            H = 2 * h;
            isClose = true;
        }
        else if (shapType === "arc") {
            adj1 = 270;
            adj2 = 0;
            H = h;
            isClose = false;
        }
        if (shapAdjst !== undefined) {
            shapAdjst1 = PPTXXmlUtils.getTextByPathList(shapAdjst, ["attrs", "fmla"]);
            shapAdjst2 = shapAdjst1;
            if (shapAdjst1 === undefined) {
                shapAdjst1 = shapAdjst[0]["attrs"]["fmla"];
                shapAdjst2 = shapAdjst[1]["attrs"]["fmla"];
            }
            if (shapAdjst1 !== undefined) {
                adj1 = parseInt(shapAdjst1.substr(4)) / 60000;
            }
            if (shapAdjst2 !== undefined) {
                adj2 = parseInt(shapAdjst2.substr(4)) / 60000;
            }
        }
        const pieVals = shapePie(H, w, adj1, adj2, isClose);
        result += `<path d='${pieVals[0]}' transform='${pieVals[1]}' fill='${(!imgFillFlg ? (grndFillFlg ? "url(#linGrd_" + shpId + ")" : fillColor) : "url(#imgPtrn_" + shpId + ")")}' stroke='${border.color}' stroke-width='${border.width}' stroke-dasharray='${border.strokeDasharray}' ${(oShadowSvgUrlStr || "")} />`;
    }
    else if (shapType === "chord") {
        const shapAdjst_ary = PPTXXmlUtils.getTextByPathList(node, ["p:spPr", "a:prstGeom", "a:avLst", "a:gd"]);
        let sAdj1, sAdj1_val = 45;
        let sAdj2, sAdj2_val = 270;
        if (shapAdjst_ary !== undefined) {
            for (const i of shapAdjst_ary.keys()) {
                const sAdj_name = PPTXXmlUtils.getTextByPathList(shapAdjst_ary[i], ["attrs", "name"]);
                if (sAdj_name === "adj1") {
                    sAdj1 = PPTXXmlUtils.getTextByPathList(shapAdjst_ary[i], ["attrs", "fmla"]);
                    sAdj1_val = parseInt(sAdj1.substr(4)) / 60000;
                }
                else if (sAdj_name === "adj2") {
                    sAdj2 = PPTXXmlUtils.getTextByPathList(shapAdjst_ary[i], ["attrs", "fmla"]);
                    sAdj2_val = parseInt(sAdj2.substr(4)) / 60000;
                }
            }
        }
        const hR = h / 2;
        const wR = w / 2;
        dVal = shapeArc(wR, hR, wR, hR, sAdj1_val, sAdj2_val, true);
        result += `<path d='${dVal}' fill='${(!imgFillFlg ? (grndFillFlg ? "url(#linGrd_" + shpId + ")" : fillColor) : "url(#imgPtrn_" + shpId + ")")}' stroke='${border.color}' stroke-width='${border.width}' stroke-dasharray='${border.strokeDasharray}' ${(oShadowSvgUrlStr || "")} />`;
    }
    else if (shapType === "blockArc") {
        const shapAdjst_ary = PPTXXmlUtils.getTextByPathList(node, ["p:spPr", "a:prstGeom", "a:avLst", "a:gd"]);
        let sAdj1, adj1 = 180;
        let sAdj2, adj2 = 0;
        let sAdj3, adj3 = 25000 * SLIDE_FACTOR$1;
        const cnstVal1 = 50000 * SLIDE_FACTOR$1;
        const cnstVal2 = 100000 * SLIDE_FACTOR$1;
        if (shapAdjst_ary !== undefined) {
            for (const i of shapAdjst_ary.keys()) {
                const sAdj_name = PPTXXmlUtils.getTextByPathList(shapAdjst_ary[i], ["attrs", "name"]);
                if (sAdj_name === "adj1") {
                    sAdj1 = PPTXXmlUtils.getTextByPathList(shapAdjst_ary[i], ["attrs", "fmla"]);
                    adj1 = parseInt(sAdj1.substr(4)) / 60000;
                }
                else if (sAdj_name === "adj2") {
                    sAdj2 = PPTXXmlUtils.getTextByPathList(shapAdjst_ary[i], ["attrs", "fmla"]);
                    adj2 = parseInt(sAdj2.substr(4)) / 60000;
                }
                else if (sAdj_name === "adj3") {
                    sAdj3 = PPTXXmlUtils.getTextByPathList(shapAdjst_ary[i], ["attrs", "fmla"]);
                    adj3 = parseInt(sAdj3.substr(4)) * SLIDE_FACTOR$1;
                }
            }
        }
        let stAng, istAng, a3, sw11, sw12, swAng, iswAng;
        const cd1 = 360;
        if (adj1 < 0)
            stAng = 0;
        else if (adj1 > cd1)
            stAng = cd1;
        else
            stAng = adj1;
        if (adj2 < 0)
            istAng = 0;
        else if (adj2 > cd1)
            istAng = cd1;
        else
            istAng = adj2;
        if (adj3 < 0)
            a3 = 0;
        else if (adj3 > cnstVal1)
            a3 = cnstVal1;
        else
            a3 = adj3;
        sw11 = istAng - stAng;
        sw12 = sw11 + cd1;
        swAng = (sw11 > 0) ? sw11 : sw12;
        iswAng = -swAng;
        const endAng = stAng + swAng;
        const iendAng = istAng + iswAng;
        let wt1, ht1, dx1, dy1, x1, y1, stRd, istRd, wd2, hd2, hc, vc;
        stRd = stAng * (Math.PI) / 180;
        istRd = istAng * (Math.PI) / 180;
        wd2 = w / 2;
        hd2 = h / 2;
        hc = w / 2;
        vc = h / 2;
        if (stAng > 90 && stAng < 270) {
            wt1 = wd2 * (Math.sin((Math.PI) / 2 - stRd));
            ht1 = hd2 * (Math.cos((Math.PI) / 2 - stRd));
            dx1 = wd2 * (Math.cos(Math.atan(ht1 / wt1)));
            dy1 = hd2 * (Math.sin(Math.atan(ht1 / wt1)));
            x1 = hc - dx1;
            y1 = vc - dy1;
        }
        else {
            wt1 = wd2 * (Math.sin(stRd));
            ht1 = hd2 * (Math.cos(stRd));
            dx1 = wd2 * (Math.cos(Math.atan(wt1 / ht1)));
            dy1 = hd2 * (Math.sin(Math.atan(wt1 / ht1)));
            x1 = hc + dx1;
            y1 = vc + dy1;
        }
        let dr, iwd2, ihd2, wt2, ht2, dx2, dy2, x2, y2;
        dr = Math.min(w, h) * a3 / cnstVal2;
        iwd2 = wd2 - dr;
        ihd2 = hd2 - dr;
        if ((endAng <= 450 && endAng > 270) || ((endAng >= 630 && endAng < 720))) {
            wt2 = iwd2 * (Math.sin(istRd));
            ht2 = ihd2 * (Math.cos(istRd));
            dx2 = iwd2 * (Math.cos(Math.atan(wt2 / ht2)));
            dy2 = ihd2 * (Math.sin(Math.atan(wt2 / ht2)));
            x2 = hc + dx2;
            y2 = vc + dy2;
        }
        else {
            wt2 = iwd2 * (Math.sin((Math.PI) / 2 - istRd));
            ht2 = ihd2 * (Math.cos((Math.PI) / 2 - istRd));
            dx2 = iwd2 * (Math.cos(Math.atan(ht2 / wt2)));
            dy2 = ihd2 * (Math.sin(Math.atan(ht2 / wt2)));
            x2 = hc - dx2;
            y2 = vc - dy2;
        }
        dVal = `M${x1},${y1}${shapeArc(wd2, hd2, wd2, hd2, stAng, endAng, false).replace("M", "L")} L${x2},${y2}${shapeArc(wd2, hd2, iwd2, ihd2, istAng, iendAng, false).replace("M", "L")} z`;
        result += `<path d='${dVal}' fill='${(!imgFillFlg ? (grndFillFlg ? "url(#linGrd_" + shpId + ")" : fillColor) : "url(#imgPtrn_" + shpId + ")")}' stroke='${border.color}' stroke-width='${border.width}' stroke-dasharray='${border.strokeDasharray}' ${(oShadowSvgUrlStr || "")} />`;
    }
    return result;
}

function renderArrow(shapType, w, h, imgFillFlg, grndFillFlg, fillColor, border, shpId, node) {
    if (["rightArrow", "leftArrow", "upArrow", "downArrow"].includes(shapType)) {
        return renderBasicArrow(shapType, w, h, imgFillFlg, grndFillFlg, fillColor, border, shpId, node);
    }
    if (["leftRightArrow", "upDownArrow"].includes(shapType)) {
        return renderDoubleArrow(shapType, w, h, imgFillFlg, grndFillFlg, fillColor, border, shpId, node);
    }
    return "";
}
function readAdjustmentParams(node, w, h) {
    const shapAdjst = PPTXXmlUtils.getTextByPathList(node, ["p:spPr", "a:prstGeom", "a:avLst", "a:gd"]);
    let sAdj1, sAdj1_val = 0.25;
    let sAdj2, sAdj2_val = 0.5;
    if (shapAdjst) {
        for (const item of shapAdjst) {
            const sAdjName = PPTXXmlUtils.getTextByPathList(item, ["attrs", "name"]);
            if (sAdjName === "adj1") {
                sAdj1 = PPTXXmlUtils.getTextByPathList(item, ["attrs", "fmla"]);
                sAdj1_val = parseInt(sAdj1.substr(4)) / 200000;
            }
            else if (sAdjName === "adj2") {
                sAdj2 = PPTXXmlUtils.getTextByPathList(item, ["attrs", "fmla"]);
                const sAdj2Val2 = parseInt(sAdj2.substr(4)) / 100000;
                const maxConst = w / h;
                sAdj2_val = sAdj2Val2 / maxConst;
            }
        }
    }
    return { sAdj1_val, sAdj2_val };
}
function renderBasicArrow(shapType, w, h, imgFillFlg, grndFillFlg, fillColor, border, shpId, node) {
    let { sAdj1_val, sAdj2_val } = readAdjustmentParams(node, w, h);
    const max_sAdj2_const = w / h;
    const shapAdjst = PPTXXmlUtils.getTextByPathList(node, ["p:spPr", "a:prstGeom", "a:avLst", "a:gd"]);
    if (shapAdjst) {
        for (const item of shapAdjst) {
            const sAdjName = PPTXXmlUtils.getTextByPathList(item, ["attrs", "name"]);
            if (sAdjName === "adj2") {
                const sAdj2 = PPTXXmlUtils.getTextByPathList(item, ["attrs", "fmla"]);
                const sAdj2Val2 = parseInt(sAdj2.substr(4)) / 100000;
                sAdj2_val = sAdj2Val2 / max_sAdj2_const;
            }
        }
    }
    let points;
    if (shapType === "rightArrow") {
        points = `${w} ${h / 2},${sAdj2_val * w} 0,${sAdj2_val * w} ${sAdj1_val * h},0 ${sAdj1_val * h},0 ${(1 - sAdj1_val) * h},${sAdj2_val * w} ${(1 - sAdj1_val) * h}, ${sAdj2_val * w} ${h}`;
    }
    else if (shapType === "leftArrow") {
        points = `0 ${h / 2},${sAdj2_val * w} ${h},${sAdj2_val * w} ${(1 - sAdj1_val) * h},${w} ${(1 - sAdj1_val) * h},${w} ${sAdj1_val * h},${sAdj2_val * w} ${sAdj1_val * h}, ${sAdj2_val * w} 0`;
    }
    else if (shapType === "upArrow") {
        const max_sAdj2_const_up = h / w;
        const { sAdj1_val: sAdj1_up} = readAdjustmentParams(node, w, h);
        const shapAdjst_up = PPTXXmlUtils.getTextByPathList(node, ["p:spPr", "a:prstGeom", "a:avLst", "a:gd"]);
        let sAdj2_val_up = 0.5;
        if (shapAdjst_up) {
            for (const item of shapAdjst_up) {
                const sAdjName = PPTXXmlUtils.getTextByPathList(item, ["attrs", "name"]);
                if (sAdjName === "adj2") {
                    const sAdj2 = PPTXXmlUtils.getTextByPathList(item, ["attrs", "fmla"]);
                    const sAdj2Val2 = parseInt(sAdj2.substr(4)) / 100000;
                    sAdj2_val_up = sAdj2Val2 / max_sAdj2_const_up;
                }
            }
        }
        points = `${w / 2} 0,0 ${sAdj2_val_up * h},${(0.5 - sAdj1_up) * w} ${sAdj2_val_up * h},${(0.5 - sAdj1_up) * w} ${h},${(0.5 + sAdj1_up) * w} ${h},${(0.5 + sAdj1_up) * w} ${sAdj2_val_up * h}, ${w} ${sAdj2_val_up * h}`;
    }
    else {
        const max_sAdj2_const_down = h / w;
        const { sAdj1_val: sAdj1_down} = readAdjustmentParams(node, w, h);
        const shapAdjst_down = PPTXXmlUtils.getTextByPathList(node, ["p:spPr", "a:prstGeom", "a:avLst", "a:gd"]);
        let sAdj2_val_down = 0.5;
        if (shapAdjst_down) {
            for (const item of shapAdjst_down) {
                const sAdjName = PPTXXmlUtils.getTextByPathList(item, ["attrs", "name"]);
                if (sAdjName === "adj2") {
                    const sAdj2 = PPTXXmlUtils.getTextByPathList(item, ["attrs", "fmla"]);
                    const sAdj2Val2 = parseInt(sAdj2.substr(4)) / 100000;
                    sAdj2_val_down = sAdj2Val2 / max_sAdj2_const_down;
                }
            }
        }
        points = `${(0.5 - sAdj1_down) * w} 0,${(0.5 - sAdj1_down) * w} ${(1 - sAdj2_val_down) * h},0 ${(1 - sAdj2_val_down) * h},${w / 2} ${h},${w} ${(1 - sAdj2_val_down) * h},${(0.5 + sAdj1_down) * w} ${(1 - sAdj2_val_down) * h}, ${(0.5 + sAdj1_down) * w} 0`;
    }
    return buildPolygon(points, imgFillFlg, grndFillFlg, fillColor, border, shpId);
}
function renderDoubleArrow(shapType, w, h, imgFillFlg, grndFillFlg, fillColor, border, shpId, node) {
    let sAdj1_val = 0.25;
    let sAdj2_val = 0.5;
    const max_sAdj2_const = w / h;
    const shapAdjst = PPTXXmlUtils.getTextByPathList(node, ["p:spPr", "a:prstGeom", "a:avLst", "a:gd"]);
    if (shapAdjst) {
        for (const item of shapAdjst) {
            const sAdjName = PPTXXmlUtils.getTextByPathList(item, ["attrs", "name"]);
            if (sAdjName === "adj1") {
                const sAdj1 = PPTXXmlUtils.getTextByPathList(item, ["attrs", "fmla"]);
                sAdj1_val = parseInt(sAdj1.substr(4)) / 200000;
            }
            else if (sAdjName === "adj2") {
                const sAdj2 = PPTXXmlUtils.getTextByPathList(item, ["attrs", "fmla"]);
                const sAdj2Val2 = parseInt(sAdj2.substr(4)) / 100000;
                sAdj2_val = sAdj2Val2 / max_sAdj2_const;
            }
        }
    }
    let points;
    if (shapType === "leftRightArrow") {
        points = `0 ${h / 2},${sAdj2_val * w} 0,${sAdj2_val * w} ${h},0 ${h},${w} ${h / 2},${sAdj2_val * w} ${w},${sAdj2_val * w} ${h},${sAdj2_val * w} 0`;
    }
    else {
        sAdj1_val = 0.25;
        sAdj2_val = 0.5;
        const max_sAdj2_const_ud = h / w;
        const shapAdjst_ud = PPTXXmlUtils.getTextByPathList(node, ["p:spPr", "a:prstGeom", "a:avLst", "a:gd"]);
        if (shapAdjst_ud) {
            for (const item of shapAdjst_ud) {
                const sAdjName = PPTXXmlUtils.getTextByPathList(item, ["attrs", "name"]);
                if (sAdjName === "adj1") {
                    const sAdj1 = PPTXXmlUtils.getTextByPathList(item, ["attrs", "fmla"]);
                    sAdj1_val = parseInt(sAdj1.substr(4)) / 200000;
                }
                else if (sAdjName === "adj2") {
                    const sAdj2 = PPTXXmlUtils.getTextByPathList(item, ["attrs", "fmla"]);
                    const sAdj2Val2 = parseInt(sAdj2.substr(4)) / 100000;
                    sAdj2_val = sAdj2Val2 / max_sAdj2_const_ud;
                }
            }
        }
        points = `${w / 2} 0,${w} ${sAdj2_val * h},${w} ${h}, ${sAdj2_val * w} ${h},${w / 2} ${h},0 ${sAdj2_val * h},0 ${sAdj2_val * h},${sAdj1_val * w} 0, ${sAdj1_val * w} 0`;
    }
    return buildPolygon(points, imgFillFlg, grndFillFlg, fillColor, border, shpId);
}
function buildPolygon(points, imgFillFlg, grndFillFlg, fillColor, border, shpId) {
    const fillUrl = !imgFillFlg
        ? (grndFillFlg ? `url(#linGrd_${shpId})` : fillColor)
        : `url(#imgPtrn_${shpId})`;
    return ` <polygon points='${points}' fill='${fillUrl}' stroke='${border.color}' stroke-width='${border.width}' stroke-dasharray='${border.strokeDasharray}' />`;
}

function renderBackPrevious(w, h, imgFillFlg, grndFillFlg, fillColor, border, shpId) {
    const hc = w / 2, vc = h / 2, ss = Math.min(w, h);
    const dx2 = ss * 3 / 8;
    const g9 = vc - dx2;
    const g10 = vc + dx2;
    const g11 = hc - dx2;
    const g12 = hc + dx2;
    const d = `M${0},${0} L${w},${0} L${w},${h} L${0},${h} zM${g11},${vc} L${g12},${g9} L${g12},${g10} z`;
    return `<path d='${d}' fill='${(!imgFillFlg ? (grndFillFlg ? "url(#linGrd_" + shpId + ")" : fillColor) : "url(#imgPtrn_" + shpId + ")")}' stroke='${border.color}' stroke-width='${border.width}' stroke-dasharray='${border.strokeDasharray}' />`;
}
function renderBeginning(w, h, imgFillFlg, grndFillFlg, fillColor, border, shpId) {
    const hc = w / 2, vc = h / 2, ss = Math.min(w, h);
    const dx2 = ss * 3 / 8;
    const g9 = vc - dx2;
    const g10 = vc + dx2;
    const g11 = hc - dx2;
    const g12 = hc + dx2;
    const g13 = ss * 3 / 4;
    const g14 = g13 / 8;
    const g15 = g13 / 4;
    const g16 = g11 + g14;
    const g17 = g11 + g15;
    const d = `M${0},${0} L${w},${0} L${w},${h} L${0},${h} zM${g17},${vc} L${g12},${g9} L${g12},${g10} zM${g16},${g9} L${g11},${g9} L${g11},${g10} L${g16},${g10} z`;
    return `<path d='${d}' fill='${(!imgFillFlg ? (grndFillFlg ? "url(#linGrd_" + shpId + ")" : fillColor) : "url(#imgPtrn_" + shpId + ")")}' stroke='${border.color}' stroke-width='${border.width}' stroke-dasharray='${border.strokeDasharray}' />`;
}
function renderDocument(w, h, imgFillFlg, grndFillFlg, fillColor, border, shpId) {
    const hc = w / 2, vc = h / 2, ss = Math.min(w, h);
    const dx2 = ss * 3 / 8;
    const g9 = vc - dx2;
    const g10 = vc + dx2;
    const dx1 = ss * 9 / 32;
    const g11 = hc - dx1;
    const g12 = hc + dx1;
    const g13 = ss * 3 / 16;
    const g14 = g12 - g13;
    const g15 = g9 + g13;
    const d = `M${0},${0} L${w},${0} L${w},${h} L${0},${h} zM${g11},${g9} L${g14},${g9} L${g12},${g15} L${g12},${g10} L${g11},${g10} zM${g14},${g9} L${g14},${g15} L${g12},${g15} z`;
    return `<path d='${d}' fill='${(!imgFillFlg ? (grndFillFlg ? "url(#linGrd_" + shpId + ")" : fillColor) : "url(#imgPtrn_" + shpId + ")")}' stroke='${border.color}' stroke-width='${border.width}' stroke-dasharray='${border.strokeDasharray}' />`;
}
function renderEnd(w, h, imgFillFlg, grndFillFlg, fillColor, border, shpId) {
    const hc = w / 2, vc = h / 2, ss = Math.min(w, h);
    const dx2 = ss * 3 / 8;
    const g9 = vc - dx2;
    const g10 = vc + dx2;
    const g11 = hc - dx2;
    const g12 = hc + dx2;
    const g13 = ss * 3 / 4;
    const g14 = g13 * 3 / 4;
    const g15 = g13 * 7 / 8;
    const g16 = g11 + g14;
    const g17 = g11 + g15;
    const d = `M${0},${h} L${w},${h} L${w},${0} L${0},${0} z M${g17},${g9} L${g12},${g9} L${g12},${g10} L${g17},${g10} z M${g16},${vc} L${g11},${g9} L${g11},${g10} z`;
    return `<path d='${d}' fill='${(!imgFillFlg ? (grndFillFlg ? "url(#linGrd_" + shpId + ")" : fillColor) : "url(#imgPtrn_" + shpId + ")")}' stroke='${border.color}' stroke-width='${border.width}' stroke-dasharray='${border.strokeDasharray}' />`;
}
function renderForwardNext(w, h, imgFillFlg, grndFillFlg, fillColor, border, shpId) {
    const hc = w / 2, vc = h / 2, ss = Math.min(w, h);
    const dx2 = ss * 3 / 8;
    const g9 = vc - dx2;
    const g10 = vc + dx2;
    const g11 = hc - dx2;
    const g12 = hc + dx2;
    const d = `M${0},${h} L${w},${h} L${w},${0} L${0},${0} z M${g12},${vc} L${g11},${g9} L${g11},${g10} z`;
    return `<path d='${d}' fill='${(!imgFillFlg ? (grndFillFlg ? "url(#linGrd_" + shpId + ")" : fillColor) : "url(#imgPtrn_" + shpId + ")")}' stroke='${border.color}' stroke-width='${border.width}' stroke-dasharray='${border.strokeDasharray}' />`;
}
function renderHelp(w, h, imgFillFlg, grndFillFlg, fillColor, border, shpId, shapeArcAlt) {
    const hc = w / 2, vc = h / 2, ss = Math.min(w, h);
    const dx2 = ss * 3 / 8;
    const g9 = vc - dx2;
    const g11 = hc - dx2;
    const g13 = ss * 3 / 4;
    const g14 = g13 / 7;
    const g15 = g13 * 3 / 14;
    const g16 = g13 * 2 / 7;
    const g19 = g13 * 3 / 7;
    const g20 = g13 * 4 / 7;
    const g21 = g13 * 17 / 28;
    const g23 = g13 * 21 / 28;
    const g24 = g13 * 11 / 14;
    const g27 = g9 + g16;
    const g29 = g9 + g21;
    const g30 = g9 + g23;
    const g31 = g9 + g24;
    const g33 = g11 + g15;
    const g36 = g11 + g19;
    const g37 = g11 + g20;
    const g41 = g13 / 14;
    const g42 = g13 * 3 / 28;
    const cX1 = g33 + g16;
    const cX2 = g36 + g14;
    const cY3 = g31 + g42;
    const cX4 = (g37 + g36 + g16) / 2;
    const d = `M${0},${0} L${w},${0} L${w},${h} L${0},${h} zM${g33},${g27}${shapeArcAlt(cX1, g27, g16, g16, 180, 360, false).replace("M", "L")}${shapeArcAlt(cX4, g27, g14, g15, 0, 90, false).replace("M", "L")}${shapeArcAlt(cX4, g29, g41, g42, 270, 180, false).replace("M", "L")} L${g37},${g30} L${g36},${g30} L${g36},${g29}${shapeArcAlt(cX2, g29, g14, g15, 180, 270, false).replace("M", "L")}${shapeArcAlt(g37, g27, g41, g42, 90, 0, false).replace("M", "L")}${shapeArcAlt(cX1, g27, g14, g14, 0, -180, false).replace("M", "L")} zM${hc},${g31}${shapeArcAlt(hc, cY3, g42, g42, 270, 630, false).replace("M", "L")} z`;
    return `<path d='${d}' fill='${(!imgFillFlg ? (grndFillFlg ? "url(#linGrd_" + shpId + ")" : fillColor) : "url(#imgPtrn_" + shpId + ")")}' stroke='${border.color}' stroke-width='${border.width}' stroke-dasharray='${border.strokeDasharray}' />`;
}
function renderHome(w, h, imgFillFlg, grndFillFlg, fillColor, border, shpId) {
    const hc = w / 2, vc = h / 2, ss = Math.min(w, h);
    const dx2 = ss * 3 / 8;
    const g9 = vc - dx2;
    const g10 = vc + dx2;
    const g11 = hc - dx2;
    const g12 = hc + dx2;
    const g13 = ss * 3 / 4;
    const g14 = g13 / 16;
    const g15 = g13 / 8;
    const g16 = g13 * 3 / 16;
    const g17 = g13 * 5 / 16;
    const g18 = g13 * 7 / 16;
    const g19 = g13 * 9 / 16;
    const g20 = g13 * 11 / 16;
    const g21 = g13 * 3 / 4;
    const g22 = g13 * 13 / 16;
    const g23 = g13 * 7 / 8;
    const g24 = g9 + g14;
    const g25 = g9 + g16;
    const g26 = g9 + g17;
    const g27 = g9 + g21;
    const g28 = g11 + g15;
    const g29 = g11 + g18;
    const g30 = g11 + g19;
    const g31 = g11 + g20;
    const g32 = g11 + g22;
    const g33 = g11 + g23;
    const d = `M${0},${0} L${w},${0} L${w},${h} L${0},${h} z M${hc},${g9} L${g11},${vc} L${g28},${vc} L${g28},${g10} L${g33},${g10} L${g33},${vc} L${g12},${vc} L${g32},${g26} L${g32},${g24} L${g31},${g24} L${g31},${g25} z M${g29},${g27} L${g30},${g27} L${g30},${g10} L${g29},${g10} z`;
    return `<path d='${d}' fill='${(!imgFillFlg ? (grndFillFlg ? "url(#linGrd_" + shpId + ")" : fillColor) : "url(#imgPtrn_" + shpId + ")")}' stroke='${border.color}' stroke-width='${border.width}' stroke-dasharray='${border.strokeDasharray}' />`;
}
function renderInformation(w, h, imgFillFlg, grndFillFlg, fillColor, border, shpId, shapeArcAlt) {
    const hc = w / 2, vc = h / 2, ss = Math.min(w, h);
    const dx2 = ss * 3 / 8;
    const g9 = vc - dx2;
    const g11 = hc - dx2;
    const g13 = ss * 3 / 4;
    const g14 = g13 / 32;
    const g17 = g13 * 5 / 16;
    const g18 = g13 * 3 / 8;
    const g19 = g13 * 13 / 32;
    const g20 = g13 * 19 / 32;
    const g22 = g13 * 11 / 16;
    const g23 = g13 * 13 / 16;
    const g24 = g13 * 7 / 8;
    const g25 = g9 + g14;
    const g28 = g9 + g17;
    const g29 = g9 + g18;
    const g30 = g9 + g23;
    const g31 = g9 + g24;
    const g32 = g11 + g17;
    const g34 = g11 + g19;
    const g35 = g11 + g20;
    const g37 = g11 + g22;
    const g38 = g13 * 3 / 32;
    const cY1 = g9 + dx2;
    const cY2 = g25 + g38;
    const d = `M${0},${0} L${w},${0} L${w},${h} L${0},${h} zM${hc},${g9}${shapeArcAlt(hc, cY1, dx2, dx2, 270, 630, false).replace("M", "L")} zM${hc},${g25}${shapeArcAlt(hc, cY2, g38, g38, 270, 630, false).replace("M", "L")}M${g32},${g28} L${g35},${g28} L${g35},${g30} L${g37},${g30} L${g37},${g31} L${g32},${g31} L${g32},${g30} L${g34},${g30} L${g34},${g29} L${g32},${g29} z`;
    return `<path d='${d}' fill='${(!imgFillFlg ? (grndFillFlg ? "url(#linGrd_" + shpId + ")" : fillColor) : "url(#imgPtrn_" + shpId + ")")}' stroke='${border.color}' stroke-width='${border.width}' stroke-dasharray='${border.strokeDasharray}' />`;
}
function renderMovie(w, h, imgFillFlg, grndFillFlg, fillColor, border, shpId) {
    const hc = w / 2, vc = h / 2, ss = Math.min(w, h);
    const dx2 = ss * 3 / 8;
    const g9 = vc - dx2;
    const g11 = hc - dx2;
    const g12 = hc + dx2;
    const g13 = ss * 3 / 4;
    const g14 = g13 * 1455 / 21600;
    const g15 = g13 * 1905 / 21600;
    const g16 = g13 * 2325 / 21600;
    const g17 = g13 * 16155 / 21600;
    const g18 = g13 * 17010 / 21600;
    const g19 = g13 * 19335 / 21600;
    const g20 = g13 * 19725 / 21600;
    const g21 = g13 * 20595 / 21600;
    const g22 = g13 * 5280 / 21600;
    const g23 = g13 * 5730 / 21600;
    const g24 = g13 * 6630 / 21600;
    const g25 = g13 * 7492 / 21600;
    const g26 = g13 * 9067 / 21600;
    const g27 = g13 * 9555 / 21600;
    const g28 = g13 * 13342 / 21600;
    const g29 = g13 * 14580 / 21600;
    const g30 = g13 * 15592 / 21600;
    const g31 = g11 + g14;
    const g32 = g11 + g15;
    const g33 = g11 + g16;
    const g34 = g11 + g17;
    const g35 = g11 + g18;
    const g36 = g11 + g19;
    const g37 = g11 + g20;
    const g38 = g11 + g21;
    const g39 = g9 + g22;
    const g40 = g9 + g23;
    const g41 = g9 + g24;
    const g42 = g9 + g25;
    const g43 = g9 + g26;
    const g44 = g9 + g27;
    const g45 = g9 + g28;
    const g46 = g9 + g29;
    const g47 = g9 + g30;
    const d = `M${0},${h} L${w},${h} L${w},${0} L${0},${0} zM${g11},${g39} L${g11},${g44} L${g31},${g44} L${g32},${g43} L${g33},${g43} L${g33},${g47} L${g35},${g47} L${g35},${g45} L${g36},${g45} L${g38},${g46} L${g12},${g46} L${g12},${g41} L${g38},${g41} L${g37},${g42} L${g35},${g42} L${g35},${g41} L${g34},${g40} L${g32},${g40} L${g31},${g39} z`;
    return `<path d='${d}' fill='${(!imgFillFlg ? (grndFillFlg ? "url(#linGrd_" + shpId + ")" : fillColor) : "url(#imgPtrn_" + shpId + ")")}' stroke='${border.color}' stroke-width='${border.width}' stroke-dasharray='${border.strokeDasharray}' />`;
}
function renderReturn(w, h, imgFillFlg, grndFillFlg, fillColor, border, shpId, shapeArcAlt) {
    const hc = w / 2, vc = h / 2, ss = Math.min(w, h);
    const dx2 = ss * 3 / 8;
    const g9 = vc - dx2;
    const g10 = vc + dx2;
    const g11 = hc - dx2;
    const g12 = hc + dx2;
    const g13 = ss * 3 / 4;
    const g14 = g13 * 7 / 8;
    const g15 = g13 * 3 / 4;
    const g16 = g13 * 5 / 8;
    const g17 = g13 * 3 / 8;
    const g18 = g13 / 4;
    const g19 = g9 + g15;
    const g20 = g9 + g16;
    const g21 = g9 + g18;
    const g22 = g11 + g14;
    const g23 = g11 + g15;
    const g24 = g11 + g16;
    const g25 = g11 + g17;
    const g26 = g11 + g18;
    const g27 = g13 / 8;
    const cX1 = g24 - g27;
    const cY2 = g19 - g27;
    const cX3 = g11 + g17;
    const cY4 = g10 - g17;
    const d = `M${0},${h} L${w},${h} L${w},${0} L${0},${0} z M${g12},${g21} L${g23},${g9} L${hc},${g21} L${g24},${g21} L${g24},${g20}${shapeArcAlt(cX1, g20, g27, g27, 0, 90, false).replace("M", "L")} L${g25},${g19}${shapeArcAlt(g25, cY2, g27, g27, 90, 180, false).replace("M", "L")} L${g26},${g21} L${g11},${g21} L${g11},${g20}${shapeArcAlt(cX3, g20, g17, g17, 180, 90, false).replace("M", "L")} L${hc},${g10}${shapeArcAlt(hc, cY4, g17, g17, 90, 0, false).replace("M", "L")} L${g22},${g21} z`;
    return `<path d='${d}' fill='${(!imgFillFlg ? (grndFillFlg ? "url(#linGrd_" + shpId + ")" : fillColor) : "url(#imgPtrn_" + shpId + ")")}' stroke='${border.color}' stroke-width='${border.width}' stroke-dasharray='${border.strokeDasharray}' />`;
}
function renderSound(w, h, imgFillFlg, grndFillFlg, fillColor, border, shpId) {
    const hc = w / 2, vc = h / 2, ss = Math.min(w, h);
    const dx2 = ss * 3 / 8;
    const g9 = vc - dx2;
    const g10 = vc + dx2;
    const g11 = hc - dx2;
    const g12 = hc + dx2;
    const g13 = ss * 3 / 4;
    const g14 = g13 / 8;
    const g15 = g13 * 5 / 16;
    const g16 = g13 * 5 / 8;
    const g17 = g13 * 11 / 16;
    const g18 = g13 * 3 / 4;
    const g19 = g13 * 7 / 8;
    const g20 = g9 + g14;
    const g21 = g9 + g15;
    const g22 = g9 + g17;
    const g23 = g9 + g19;
    const g24 = g11 + g15;
    const g25 = g11 + g16;
    const g26 = g11 + g18;
    const d = `M${0},${0} L${w},${0} L${w},${h} L${0},${h} z M${g11},${g21} L${g24},${g21} L${g25},${g9} L${g25},${g10} L${g24},${g22} L${g11},${g22} z M${g26},${g21} L${g12},${g20} M${g26},${vc} L${g12},${vc} M${g26},${g22} L${g12},${g23}`;
    return `<path d='${d}' fill='${(!imgFillFlg ? (grndFillFlg ? "url(#linGrd_" + shpId + ")" : fillColor) : "url(#imgPtrn_" + shpId + ")")}' stroke='${border.color}' stroke-width='${border.width}' stroke-dasharray='${border.strokeDasharray}' />`;
}
const BUTTON_RENDERERS = {
    'actionButtonBackPrevious': renderBackPrevious,
    'actionButtonBeginning': renderBeginning,
    'actionButtonDocument': renderDocument,
    'actionButtonEnd': renderEnd,
    'actionButtonForwardNext': renderForwardNext,
    'actionButtonHelp': renderHelp,
    'actionButtonHome': renderHome,
    'actionButtonInformation': renderInformation,
    'actionButtonMovie': renderMovie,
    'actionButtonReturn': renderReturn,
    'actionButtonSound': renderSound
};
function renderActionButton(shapeType, w, h, imgFillFlg, grndFillFlg, fillColor, border, shpId, shapeArcAlt) {
    const renderer = BUTTON_RENDERERS[shapeType];
    if (!renderer) {
        return '';
    }
    return renderer(w, h, imgFillFlg, grndFillFlg, fillColor, border, shpId, shapeArcAlt);
}

const PPTXShapeUtils = (function () {
    function genShapeDataAttributes(node, slideXfrmNode, id, name, idx, type, rotate, sType) {
        let dataAttrs = '';
        let offX = 0, offY = 0, extCx = 0, extCy = 0, flipH = 0, flipV = 0;
        if (slideXfrmNode !== undefined) {
            if (slideXfrmNode['a:off'] && slideXfrmNode['a:off'].attrs) {
                offX = slideXfrmNode['a:off'].attrs.x || 0;
                offY = slideXfrmNode['a:off'].attrs.y || 0;
            }
            if (slideXfrmNode['a:ext'] && slideXfrmNode['a:ext'].attrs) {
                extCx = slideXfrmNode['a:ext'].attrs.cx || 0;
                extCy = slideXfrmNode['a:ext'].attrs.cy || 0;
            }
            if (slideXfrmNode['attrs']) {
                slideXfrmNode['attrs'].rot || 0;
                flipH = slideXfrmNode['attrs'].flipH || '0';
                flipV = slideXfrmNode['attrs'].flipV || '0';
            }
        }
        const shapType = PPTXXmlUtils.getTextByPathList(node, ["p:spPr", "a:prstGeom", "attrs", "prst"]);
        dataAttrs += ` data-node-id="${id || ''}"`;
        dataAttrs += ` data-node-name="${name || ''}"`;
        dataAttrs += ` data-node-idx="${idx || ''}"`;
        dataAttrs += ` data-node-type="${type || ''}"`;
        dataAttrs += ` data-shape-type="${sType || ''}"`;
        dataAttrs += ` data-off-x="${offX}"`;
        dataAttrs += ` data-off-y="${offY}"`;
        dataAttrs += ` data-ext-cx="${extCx}"`;
        dataAttrs += ` data-ext-cy="${extCy}"`;
        dataAttrs += ` data-rotate="${rotate || 0}"`;
        dataAttrs += ` data-flip-h="${flipH}"`;
        dataAttrs += ` data-flip-v="${flipV}"`;
        if (shapType) {
            dataAttrs += ` data-geom-type="${shapType}"`;
        }
        return dataAttrs;
    }
    async function genShape(node, pNode, slideLayoutSpNode, slideMasterSpNode, id, name, idx, type, order, warpObj, isUserDrawnBg, sType, source, settings) {
        const xfrmList = ["p:spPr", "a:xfrm"];
        const slideXfrmNode = PPTXXmlUtils.getTextByPathList(node, xfrmList);
        const slideLayoutXfrmNode = PPTXXmlUtils.getTextByPathList(slideLayoutSpNode, xfrmList);
        const slideMasterXfrmNode = PPTXXmlUtils.getTextByPathList(slideMasterSpNode, xfrmList);
        let result = "";
        const shpId = PPTXXmlUtils.getTextByPathList(node, ["attrs", "order"]);
        const shapType = PPTXXmlUtils.getTextByPathList(node, ["p:spPr", "a:prstGeom", "attrs", "prst"]);
        let transform3dStyle = "";
        const custShapType = PPTXXmlUtils.getTextByPathList(node, ["p:spPr", "a:custGeom"]);
        let isFlipV = false;
        let isFlipH = false;
        let flip = "";
        const flipVAttr = PPTXXmlUtils.getTextByPathList(slideXfrmNode, ["attrs", "flipV"]);
        const flipHAttr = PPTXXmlUtils.getTextByPathList(slideXfrmNode, ["attrs", "flipH"]);
        if (flipVAttr === "1" || flipVAttr === "true") {
            isFlipV = true;
        }
        if (flipHAttr === "1" || flipHAttr === "true") {
            isFlipH = true;
        }
        if (isFlipH && !isFlipV) {
            flip = " scale(-1,1)";
        }
        else if (!isFlipH && isFlipV) {
            flip = " scale(1,-1)";
        }
        else if (isFlipH && isFlipV) {
            flip = " scale(-1,-1)";
        }
        const rotate = PPTXXmlUtils.angleToDegrees(PPTXXmlUtils.getTextByPathList(slideXfrmNode, ["attrs", "rot"]));
        const txtXframeNode = PPTXXmlUtils.getTextByPathList(node, ["p:txXfrm"]);
        if (txtXframeNode !== undefined) {
            PPTXXmlUtils.getTextByPathList(txtXframeNode, ["attrs", "rot"]);
        }
        let workingXfrmNode = slideXfrmNode;
        let drawW, drawH;
        if (slideXfrmNode && slideXfrmNode['a:ext'] && slideXfrmNode['a:ext'].attrs) {
            const originalCx = parseInt(slideXfrmNode['a:ext'].attrs.cx);
            const originalCy = parseInt(slideXfrmNode['a:ext'].attrs.cy);
            drawW = (originalCx !== undefined && originalCy !== undefined)
                ? originalCx * SLIDE_FACTOR$1
                : undefined;
            drawH = (originalCx !== undefined && originalCy !== undefined)
                ? originalCy * SLIDE_FACTOR$1
                : undefined;
        }
        if (sType === 'group-abs' && warpObj.currentGroupScale && slideXfrmNode) {
            const { scaleX, scaleY, childX, childY } = warpObj.currentGroupScale;
            workingXfrmNode = JSON.parse(JSON.stringify(slideXfrmNode));
            if (slideXfrmNode['a:ext'] && slideXfrmNode['a:ext'].attrs) {
                const originalCx = parseInt(slideXfrmNode['a:ext'].attrs.cx);
                const originalCy = parseInt(slideXfrmNode['a:ext'].attrs.cy);
                workingXfrmNode['a:ext'].attrs.cx = Math.round(originalCx * scaleX);
                workingXfrmNode['a:ext'].attrs.cy = Math.round(originalCy * scaleY);
            }
            if (slideXfrmNode['a:off'] && slideXfrmNode['a:off'].attrs) {
                const originalOffX = parseInt(slideXfrmNode['a:off'].attrs.x);
                const originalOffY = parseInt(slideXfrmNode['a:off'].attrs.y);
                const childXEmu = childX / SLIDE_FACTOR$1;
                const childYEmu = childY / SLIDE_FACTOR$1;
                const relativeX = originalOffX - childXEmu;
                const relativeY = originalOffY - childYEmu;
                workingXfrmNode['a:off'].attrs.x = Math.round(childXEmu + relativeX * scaleX);
                workingXfrmNode['a:off'].attrs.y = Math.round(childYEmu + relativeY * scaleY);
            }
        }
        if (shapType !== undefined || custShapType !== undefined) {
            const off = PPTXXmlUtils.getTextByPathList(workingXfrmNode, ["a:off", "attrs"]);
            (off !== undefined) ? parseInt(off["x"]) * SLIDE_FACTOR$1 : 0;
            (off !== undefined) ? parseInt(off["y"]) * SLIDE_FACTOR$1 : 0;
            let ext = PPTXXmlUtils.getTextByPathList(workingXfrmNode, ["a:ext", "attrs"]);
            if (ext === undefined && slideLayoutXfrmNode !== undefined) {
                ext = PPTXXmlUtils.getTextByPathList(slideLayoutXfrmNode, ["a:ext", "attrs"]);
            }
            if (ext === undefined && slideMasterXfrmNode !== undefined) {
                ext = PPTXXmlUtils.getTextByPathList(slideMasterXfrmNode, ["a:ext", "attrs"]);
            }
            var w = (ext !== undefined && ext["cx"] !== undefined) ? parseInt(ext["cx"]) * SLIDE_FACTOR$1 : 100;
            var h = (ext !== undefined && ext["cy"] !== undefined) ? parseInt(ext["cy"]) * SLIDE_FACTOR$1 : 100;
            w = isNaN(w) ? 100 : w;
            h = isNaN(h) ? 100 : h;
            if (drawW === undefined)
                drawW = w;
            if (drawH === undefined)
                drawH = h;
            const isConnector = (shapType === 'straightConnector1' || shapType === 'bentConnector2' ||
                shapType === 'bentConnector3' || shapType === 'bentConnector4' ||
                shapType === 'bentConnector5' || shapType === 'curvedConnector2' ||
                shapType === 'curvedConnector3' || shapType === 'curvedConnector4' ||
                shapType === 'curvedConnector5');
            const svgCssName = `_svg_css_${(Object.keys(warpObj.styleTable).length + 1)}_${Math.floor(Math.random() * 1001)}`;
            const effectsClassName = `${svgCssName}_effects`;
            let svgSizeStyle = "";
            if (isConnector && (w === 0 || h === 0)) {
                const strokeWidth = 1.5;
                const minSize = Math.max(strokeWidth * 2, 4);
                const svgW = (w === 0 || w < minSize) ? minSize : w;
                const svgH = (h === 0 || h < minSize) ? minSize : h;
                svgSizeStyle = `width:${svgW}px; height:${svgH}px; overflow: visible;`;
                w = svgW;
                h = svgH;
            }
            else {
                svgSizeStyle = `${PPTXXmlUtils.getSize(workingXfrmNode, undefined, undefined)} overflow: visible;`;
            }
            let svgTransform = `transform: rotate(${((rotate !== undefined) ? rotate : 0)}deg)${flip};`;
            if (sType === 'group-abs' && warpObj.currentGroupScale) {
                if (custShapType === undefined) {
                    const { scaleX, scaleY } = warpObj.currentGroupScale;
                    svgTransform = `transform: rotate(${(rotate !== undefined) ? rotate : 0}deg)${flip} scale(${scaleX},${scaleY});`;
                }
            }
            const svgTag = `<svg class='drawing ${svgCssName}' _id='${id}' _idx='${idx}' _type='${type}' _name='${name}'' style='${PPTXXmlUtils.getPosition(workingXfrmNode, pNode, undefined, undefined, sType)}${svgSizeStyle} z-index: ${order};${svgTransform}'>`;
            result += svgTag;
            result += '<defs>';
            var fillColor = await PPTXStyleUtils.getShapeFill(node, pNode, true, warpObj, source);
            var grndFillFlg = false;
            var imgFillFlg = false;
            let clrFillType = PPTXStyleUtils.getFillType(PPTXXmlUtils.getTextByPathList(node, ["p:spPr"]));
            if (clrFillType == "GROUP_FILL") {
                clrFillType = PPTXStyleUtils.getFillType(PPTXXmlUtils.getTextByPathList(pNode, ["p:grpSpPr"]));
            }
            if (clrFillType == "GRADIENT_FILL") {
                grndFillFlg = true;
                const color_arry = fillColor.color;
                const angl = fillColor.rot + 90;
                const svgGrdnt = PPTXStyleUtils.getSvgGradient(w, h, angl, color_arry, shpId);
                result += svgGrdnt;
            }
            else if (clrFillType == "PIC_FILL") {
                imgFillFlg = true;
                const svgBgImg = PPTXStyleUtils.getSvgImagePattern(node, fillColor, shpId, warpObj);
                result += svgBgImg;
            }
            else if (clrFillType == "PATTERN_FILL") {
                let styleText = fillColor;
                if (styleText in warpObj.styleTable) {
                    styleText += `do-nothing: ${svgCssName};`;
                }
                warpObj.styleTable[styleText] = {
                    "name": svgCssName,
                    "text": styleText
                };
                fillColor = "none";
            }
            else {
                if (clrFillType != "SOLID_FILL" && clrFillType != "PATTERN_FILL" &&
                    (shapType == "arc" ||
                        shapType == "bracketPair" ||
                        shapType == "bracePair" ||
                        shapType == "leftBracket" ||
                        shapType == "leftBrace" ||
                        shapType == "rightBrace" ||
                        shapType == "rightBracket")) {
                    fillColor = "none";
                }
            }
            var border = PPTXStyleUtils.getBorder(node, pNode, true, "shape", warpObj);
            var headEndNodeAttrs = PPTXXmlUtils.getTextByPathList(node, ["p:spPr", "a:ln", "a:headEnd", "attrs"]);
            var tailEndNodeAttrs = PPTXXmlUtils.getTextByPathList(node, ["p:spPr", "a:ln", "a:tailEnd", "attrs"]);
            const scene3d = PPTXXmlUtils.getTextByPathList(node, ["p:spPr", "a:scene3d"]);
            const sp3d = PPTXXmlUtils.getTextByPathList(node, ["p:spPr", "a:sp3d"]);
            if (scene3d || sp3d) {
                transform3dStyle = process3DEffects(scene3d, sp3d);
            }
            const effectRefNode = PPTXXmlUtils.getTextByPathList(node, ["p:style", "a:effectRef"]);
            let effectStyleNode = undefined;
            if (effectRefNode !== undefined) {
                const effectIdx = PPTXXmlUtils.getTextByPathList(effectRefNode, ["attrs", "idx"]);
                if (effectIdx !== undefined && warpObj["themeContent"] !== undefined) {
                    let effectStyleLst = warpObj["themeContent"]["a:theme"]["a:themeElements"]["a:fmtScheme"]["a:effectStyleLst"]["a:effectStyle"];
                    if (effectStyleLst !== undefined) {
                        if (!Array.isArray(effectStyleLst)) {
                            effectStyleLst = [effectStyleLst];
                        }
                        var idx = Number(effectIdx);
                        if (effectStyleLst.length > 0) {
                            if (idx >= 0 && idx < effectStyleLst.length) {
                                effectStyleNode = effectStyleLst[idx];
                            }
                            else {
                                for (var i = effectStyleLst.length - 1; i >= 0; i--) {
                                    const testEffectStyle = effectStyleLst[i];
                                    const hasShadow = PPTXXmlUtils.getTextByPathList(testEffectStyle, ["a:effectLst", "a:outerShdw"]);
                                    if (hasShadow !== undefined) {
                                        effectStyleNode = testEffectStyle;
                                        break;
                                    }
                                }
                                if (effectStyleNode === undefined) {
                                    idx = idx % effectStyleLst.length;
                                    if (idx < 0)
                                        idx += effectStyleLst.length;
                                    effectStyleNode = effectStyleLst[idx];
                                }
                            }
                        }
                    }
                }
            }
            let outerShdwNode = PPTXXmlUtils.getTextByPathList(node, ["p:spPr", "a:effectLst", "a:outerShdw"]);
            if (outerShdwNode === undefined && effectStyleNode !== undefined) {
                outerShdwNode = PPTXXmlUtils.getTextByPathList(effectStyleNode, ["a:effectLst", "a:outerShdw"]);
            }
            var oShadowSvgUrlStr = "";
            let hasOuterShadow = false;
            if (outerShdwNode && typeof outerShdwNode === 'object' && !Array.isArray(outerShdwNode)) {
                const nodeKeys = Object.keys(outerShdwNode);
                if (nodeKeys.length > 0) {
                    var attrs = outerShdwNode.attrs;
                    if (attrs && typeof attrs === 'object') {
                        const { dist: distVal, blurRad: blurRadVal } = attrs;
                        const hasShadowAttrs = (distVal !== undefined || blurRadVal !== undefined ||
                            attrs.dir !== undefined || attrs.sx !== undefined ||
                            attrs.sy !== undefined || attrs.algn !== undefined);
                        hasOuterShadow = hasShadowAttrs && (distVal !== undefined && distVal !== "" && distVal !== "0" && distVal !== 0);
                    }
                }
            }
            const sp3dNode = PPTXXmlUtils.getTextByPathList(node, ["p:spPr", "a:sp3d"]);
            const scene3dNode = PPTXXmlUtils.getTextByPathList(node, ["p:spPr", "a:scene3d"]);
            const shadowFromEffectStyle = (outerShdwNode !== undefined && effectStyleNode !== undefined);
            if ((sp3dNode !== undefined || scene3dNode !== undefined) && !shadowFromEffectStyle) {
                hasOuterShadow = false;
            }
            if (hasOuterShadow) {
                const chdwClrNode = PPTXStyleUtils.getSolidFill(outerShdwNode, undefined, undefined, warpObj);
                const outerShdwAttrs = outerShdwNode["attrs"];
                var dir = (outerShdwAttrs["dir"]) ? (parseInt(outerShdwAttrs["dir"]) / 60000) : 0;
                var dist = parseInt(outerShdwAttrs["dist"]) * SLIDE_FACTOR$1;
                var blurRad = (outerShdwAttrs["blurRad"]) ? (parseInt(outerShdwAttrs["blurRad"]) * SLIDE_FACTOR$1) : "";
                const vx = dist * Math.sin(dir * Math.PI / 180);
                const hx = dist * Math.cos(dir * Math.PI / 180);
                let svg_css_shadow = `filter:drop-shadow(${hx}px ${vx}px ${blurRad}px #${chdwClrNode});`;
                if (svg_css_shadow in warpObj.styleTable) {
                    svg_css_shadow += `do-nothing: ${svgCssName};`;
                }
                warpObj.styleTable[svg_css_shadow] = {
                    "name": effectsClassName,
                    "text": svg_css_shadow
                };
                result = result.replace(`class='drawing ${svgCssName}'`, `class='drawing ${svgCssName} ${effectsClassName}'`);
            }
            let softEdgeNode = PPTXXmlUtils.getTextByPathList(node, ["p:spPr", "a:effectLst", "a:softEdge"]);
            if (softEdgeNode === undefined && effectStyleNode !== undefined) {
                softEdgeNode = PPTXXmlUtils.getTextByPathList(effectStyleNode, ["a:effectLst", "a:softEdge"]);
            }
            var softEdgeFilterStr = "";
            if (softEdgeNode !== undefined) {
                const softEdgeAttrs = softEdgeNode["attrs"];
                const rad = (softEdgeAttrs["rad"]) ? (parseInt(softEdgeAttrs["rad"]) * SLIDE_FACTOR$1) : 0;
                const softEdgeId = `softedge_${shpId}`;
                let softEdgeFilter = `<filter id="${softEdgeId}" x="-20%" y="-20%" width="140%" height="140%">`;
                softEdgeFilter += `<feGaussianBlur in="SourceGraphic" stdDeviation="${rad}" />`;
                softEdgeFilter += '</filter>';
                result += softEdgeFilter;
                softEdgeFilterStr = `filter="url(#${softEdgeId})"`;
            }
            if ((headEndNodeAttrs !== undefined && (headEndNodeAttrs["type"] === "triangle" || headEndNodeAttrs["type"] === "arrow")) ||
                (tailEndNodeAttrs !== undefined && (tailEndNodeAttrs["type"] === "triangle" || tailEndNodeAttrs["type"] === "arrow"))) {
                const triangleMarker = `<marker id='markerTriangle_${shpId}' viewBox='0 0 10 10' refX='10' refY='5' markerWidth='5' markerHeight='5' stroke='${border.color}' fill='${border.color}' orient='auto-start-reverse' markerUnits='strokeWidth'><path d='M 0 0 L 10 5 L 0 10 z' /></marker>`;
                result += triangleMarker;
            }
            result += '</defs>';
        }
        if (shapType !== undefined && custShapType === undefined) {
            switch (shapType) {
                case "rect":
                case "flowChartProcess":
                case "flowChartPredefinedProcess":
                case "flowChartInternalStorage":
                case "actionButtonBlank": {
                    result += `<rect x='0' y='0' width='${w}' height='${h}' fill='${(!imgFillFlg ? (grndFillFlg ? "url(#linGrd_" + shpId + ")" : fillColor) : "url(#imgPtrn_" + shpId + ")")}' stroke='${border.color}' stroke-width='${border.width}' stroke-dasharray='${border.strokeDasharray}' ${oShadowSvgUrlStr}  />`;
                    if (shapType == "flowChartPredefinedProcess") {
                        result += `<rect x='${w * (1 / 8)}' y='0' width='${w * (6 / 8)}' height='${h}' fill='none' stroke='${border.color}' stroke-width='${border.width}' stroke-dasharray='${border.strokeDasharray}' />`;
                    }
                    else if (shapType == "flowChartInternalStorage") {
                        result += ` <polyline points='${w * (1 / 8)} 0,${w * (1 / 8)} ${h}' fill='none' stroke='${border.color}' stroke-width='${border.width}' stroke-dasharray='${border.strokeDasharray}' />`;
                        result += ` <polyline points='0 ${h * (1 / 8)},${w} ${h * (1 / 8)}' fill='none' stroke='${border.color}' stroke-width='${border.width}' stroke-dasharray='${border.strokeDasharray}' />`;
                    }
                    break;
                }
                case "flowChartCollate": {
                    var d = `M 0,0 L${w},${0} L${0},${h} L${w},${h} z`;
                    result += `<path d='${d}'  fill='${(!imgFillFlg ? (grndFillFlg ? "url(#linGrd_" + shpId + ")" : fillColor) : "url(#imgPtrn_" + shpId + ")")}' stroke='${border.color}' stroke-width='${border.width}' stroke-dasharray='${border.strokeDasharray}' ${oShadowSvgUrlStr} ${softEdgeFilterStr} />`;
                    break;
                }
                case "flowChartDocument": {
                    var y1, y2, y3, x1;
                    x1 = w * 10800 / 21600;
                    y1 = h * 17322 / 21600;
                    y2 = h * 20172 / 21600;
                    y3 = h * 23922 / 21600;
                    var d = `M${0},${0} L${w},${0} L${w},${y1} C${x1},${y1} ${x1},${y3} ${0},${y2} z`;
                    result += `<path d='${d}'  fill='${(!imgFillFlg ? (grndFillFlg ? "url(#linGrd_" + shpId + ")" : fillColor) : "url(#imgPtrn_" + shpId + ")")}' stroke='${border.color}' stroke-width='${border.width}' stroke-dasharray='${border.strokeDasharray}' />`;
                    break;
                }
                case "flowChartMultidocument": {
                    var y1, y2, y3, y4, y5, y6, y7, y8, y9, x1, x2, x3, x4, x5, x6, x7;
                    y1 = h * 18022 / 21600;
                    y2 = h * 3675 / 21600;
                    y3 = h * 23542 / 21600;
                    y4 = h * 1815 / 21600;
                    y5 = h * 16252 / 21600;
                    y6 = h * 16352 / 21600;
                    y7 = h * 14392 / 21600;
                    y8 = h * 20782 / 21600;
                    y9 = h * 14467 / 21600;
                    x1 = w * 1532 / 21600;
                    x2 = w * 20000 / 21600;
                    x3 = w * 9298 / 21600;
                    x4 = w * 19298 / 21600;
                    x5 = w * 18595 / 21600;
                    x6 = w * 2972 / 21600;
                    x7 = w * 20800 / 21600;
                    var d = `M${0},${y2} L${x5},${y2} L${x5},${y1} C${x3},${y1} ${x3},${y3} ${0},${y8} zM${x1},${y2} L${x1},${y4} L${x2},${y4} L${x2},${y5} C${x4},${y5} ${x5},${y6} ${x5},${y6}M${x6},${y4} L${x6},${0} L${w},${0} L${w},${y7} C${x7},${y7} ${x2},${y9} ${x2},${y9}`;
                    result += `<path d='${d}'  fill='${(!imgFillFlg ? (grndFillFlg ? "url(#linGrd_" + shpId + ")" : fillColor) : "url(#imgPtrn_" + shpId + ")")}' stroke='${border.color}' stroke-width='${border.width}' stroke-dasharray='${border.strokeDasharray}' />`;
                    break;
                }
                case "actionButtonBackPrevious":
                case "actionButtonBeginning":
                case "actionButtonDocument":
                case "actionButtonEnd":
                case "actionButtonForwardNext":
                case "actionButtonHelp":
                case "actionButtonHome":
                case "actionButtonInformation":
                case "actionButtonMovie":
                case "actionButtonReturn":
                case "actionButtonSound": {
                    result += renderActionButton(shapType, w, h, imgFillFlg, grndFillFlg, fillColor, border, shpId, shapeArcAlt);
                    break;
                }
                case "irregularSeal1":
                case "irregularSeal2": {
                    if (shapType == "irregularSeal1") {
                        var d = `M${w * 10800 / 21600},${h * 5800 / 21600} L${w * 14522 / 21600},${0} L${w * 14155 / 21600},${h * 5325 / 21600} L${w * 18380 / 21600},${h * 4457 / 21600} L${w * 16702 / 21600},${h * 7315 / 21600} L${w * 21097 / 21600},${h * 8137 / 21600} L${w * 17607 / 21600},${h * 10475 / 21600} L${w},${h * 13290 / 21600} L${w * 16837 / 21600},${h * 12942 / 21600} L${w * 18145 / 21600},${h * 18095 / 21600} L${w * 14020 / 21600},${h * 14457 / 21600} L${w * 13247 / 21600},${h * 19737 / 21600} L${w * 10532 / 21600},${h * 14935 / 21600} L${w * 8485 / 21600},${h} L${w * 7715 / 21600},${h * 15627 / 21600} L${w * 4762 / 21600},${h * 17617 / 21600} L${w * 5667 / 21600},${h * 13937 / 21600} L${w * 135 / 21600},${h * 14587 / 21600} L${w * 3722 / 21600},${h * 11775 / 21600} L${0},${h * 8615 / 21600} L${w * 4627 / 21600},${h * 7617 / 21600} L${w * 370 / 21600},${h * 2295 / 21600} L${w * 7312 / 21600},${h * 6320 / 21600} L${w * 8352 / 21600},${h * 2295 / 21600} z`;
                    }
                    else if (shapType == "irregularSeal2") {
                        var d = `M${w * 11462 / 21600},${h * 4342 / 21600} L${w * 14790 / 21600},${0} L${w * 14525 / 21600},${h * 5777 / 21600} L${w * 18007 / 21600},${h * 3172 / 21600} L${w * 16380 / 21600},${h * 6532 / 21600} L${w},${h * 6645 / 21600} L${w * 16985 / 21600},${h * 9402 / 21600} L${w * 18270 / 21600},${h * 11290 / 21600} L${w * 16380 / 21600},${h * 12310 / 21600} L${w * 18877 / 21600},${h * 15632 / 21600} L${w * 14640 / 21600},${h * 14350 / 21600} L${w * 14942 / 21600},${h * 17370 / 21600} L${w * 12180 / 21600},${h * 15935 / 21600} L${w * 11612 / 21600},${h * 18842 / 21600} L${w * 9872 / 21600},${h * 17370 / 21600} L${w * 8700 / 21600},${h * 19712 / 21600} L${w * 7527 / 21600},${h * 18125 / 21600} L${w * 4917 / 21600},${h} L${w * 4805 / 21600},${h * 18240 / 21600} L${w * 1285 / 21600},${h * 17825 / 21600} L${w * 3330 / 21600},${h * 15370 / 21600} L${0},${h * 12877 / 21600} L${w * 3935 / 21600},${h * 11592 / 21600} L${w * 1172 / 21600},${h * 8270 / 21600} L${w * 5372 / 21600},${h * 7817 / 21600} L${w * 4502 / 21600},${h * 3625 / 21600} L${w * 8550 / 21600},${h * 6382 / 21600} L${w * 9722 / 21600},${h * 1887 / 21600} z`;
                    }
                    result += `<path d='${d}'  fill='${(!imgFillFlg ? (grndFillFlg ? "url(#linGrd_" + shpId + ")" : fillColor) : "url(#imgPtrn_" + shpId + ")")}' stroke='${border.color}' stroke-width='${border.width}' stroke-dasharray='${border.strokeDasharray}' />`;
                    break;
                }
                case "flowChartTerminator": {
                    var x1, x2, y1, cd2 = 180, cd4 = 90, c3d4 = 270;
                    x1 = w * 3475 / 21600;
                    x2 = w * 18125 / 21600;
                    y1 = h * 10800 / 21600;
                    var d = `M${x1},${0} L${x2},${0}${PPTXShapeUtils.shapeArcAlt(x2, h / 2, x1, y1, c3d4, c3d4 + cd2, false).replace("M", "L")} L${x1},${h}${PPTXShapeUtils.shapeArcAlt(x1, h / 2, x1, y1, cd4, cd4 + cd2, false).replace("M", "L")} z`;
                    result += `<path d='${d}'  fill='${(!imgFillFlg ? (grndFillFlg ? "url(#linGrd_" + shpId + ")" : fillColor) : "url(#imgPtrn_" + shpId + ")")}' stroke='${border.color}' stroke-width='${border.width}' stroke-dasharray='${border.strokeDasharray}' />`;
                    break;
                }
                case "flowChartPunchedTape": {
                    var x1, x1, y1, y2, cd2 = 180;
                    x1 = w * 5 / 20;
                    y1 = h * 2 / 20;
                    y2 = h * 18 / 20;
                    var d = `M${0},${y1}${PPTXShapeUtils.shapeArcAlt(x1, y1, x1, y1, cd2, 0, false).replace("M", "L")}${PPTXShapeUtils.shapeArcAlt(w * (3 / 4), y1, x1, y1, cd2, 360, false).replace("M", "L")} L${w},${y2}${PPTXShapeUtils.shapeArcAlt(w * (3 / 4), y2, x1, y1, 0, -cd2, false).replace("M", "L")}${PPTXShapeUtils.shapeArcAlt(x1, y2, x1, y1, 0, cd2, false).replace("M", "L")} z`;
                    result += `<path d='${d}'  fill='${(!imgFillFlg ? (grndFillFlg ? "url(#linGrd_" + shpId + ")" : fillColor) : "url(#imgPtrn_" + shpId + ")")}' stroke='${border.color}' stroke-width='${border.width}' stroke-dasharray='${border.strokeDasharray}' />`;
                    break;
                }
                case "flowChartOnlineStorage": {
                    var x1, y1, c3d4 = 270, cd4 = 90;
                    x1 = w * 1 / 6;
                    y1 = h * 3 / 6;
                    var d = `M${x1},${0} L${w},${0}${PPTXShapeUtils.shapeArcAlt(w, h / 2, x1, y1, c3d4, 90, false).replace("M", "L")} L${x1},${h}${PPTXShapeUtils.shapeArcAlt(x1, h / 2, x1, y1, cd4, 270, false).replace("M", "L")} z`;
                    result += `<path d='${d}'  fill='${(!imgFillFlg ? (grndFillFlg ? "url(#linGrd_" + shpId + ")" : fillColor) : "url(#imgPtrn_" + shpId + ")")}' stroke='${border.color}' stroke-width='${border.width}' stroke-dasharray='${border.strokeDasharray}' />`;
                    break;
                }
                case "flowChartDisplay": {
                    var x1, x2, y1, c3d4 = 270, cd2 = 180;
                    x1 = w * 1 / 6;
                    x2 = w * 5 / 6;
                    y1 = h * 3 / 6;
                    var d = `M${0},${y1} L${x1},${0} L${x2},${0}${PPTXShapeUtils.shapeArcAlt(w, h / 2, x1, y1, c3d4, c3d4 + cd2, false).replace("M", "L")} L${x1},${h} z`;
                    result += `<path d='${d}'  fill='${(!imgFillFlg ? (grndFillFlg ? "url(#linGrd_" + shpId + ")" : fillColor) : "url(#imgPtrn_" + shpId + ")")}' stroke='${border.color}' stroke-width='${border.width}' stroke-dasharray='${border.strokeDasharray}' />`;
                    break;
                }
                case "flowChartDelay": {
                    var wd2 = w / 2, hd2 = h / 2, cd2 = 180, c3d4 = 270, cd4 = 90;
                    var d = `M${0},${0} L${wd2},${0}${PPTXShapeUtils.shapeArc(wd2, hd2, wd2, hd2, c3d4, c3d4 + cd2, false).replace("M", "L")} L${0},${h} z`;
                    result += `<path d='${d}'  fill='${(!imgFillFlg ? (grndFillFlg ? "url(#linGrd_" + shpId + ")" : fillColor) : "url(#imgPtrn_" + shpId + ")")}' stroke='${border.color}' stroke-width='${border.width}' stroke-dasharray='${border.strokeDasharray}' />`;
                    break;
                }
                case "flowChartMagneticTape": {
                    var wd2 = w / 2, hd2 = h / 2, cd2 = 180, c3d4 = 270, cd4 = 90;
                    let idy, ib, ang1;
                    idy = hd2 * Math.sin(Math.PI / 4);
                    ib = hd2 + idy;
                    ang1 = Math.atan(h / w);
                    const ang1Dg = ang1 * 180 / Math.PI;
                    var d = `M${wd2},${h}${PPTXShapeUtils.shapeArcAlt(wd2, hd2, wd2, hd2, cd4, cd2, false).replace("M", "L")}${PPTXShapeUtils.shapeArcAlt(wd2, hd2, wd2, hd2, cd2, c3d4, false).replace("M", "L")}${PPTXShapeUtils.shapeArcAlt(wd2, hd2, wd2, hd2, c3d4, 360, false).replace("M", "L")}${PPTXShapeUtils.shapeArcAlt(wd2, hd2, wd2, hd2, 0, ang1Dg, false).replace("M", "L")} L${w},${ib} L${w},${h} z`;
                    result += `<path d='${d}'  fill='${(!imgFillFlg ? (grndFillFlg ? "url(#linGrd_" + shpId + ")" : fillColor) : "url(#imgPtrn_" + shpId + ")")}' stroke='${border.color}' stroke-width='${border.width}' stroke-dasharray='${border.strokeDasharray}' />`;
                    break;
                }
                case "ellipse":
                case "flowChartConnector":
                case "flowChartSummingJunction":
                case "flowChartOr": {
                    result += `<ellipse cx='${(w / 2)}' cy='${(h / 2)}' rx='${(w / 2)}' ry='${(h / 2)}' fill='${(!imgFillFlg ? (grndFillFlg ? "url(#linGrd_" + shpId + ")" : fillColor) : "url(#imgPtrn_" + shpId + ")")}' stroke='${border.color}' stroke-width='${border.width}' stroke-dasharray='${border.strokeDasharray}' />`;
                    if (shapType == "flowChartOr") {
                        result += ` <polyline points='${w / 2} ${0},${w / 2} ${h}' fill='none' stroke='${border.color}' stroke-width='${border.width}' stroke-dasharray='${border.strokeDasharray}' />`;
                        result += ` <polyline points='${0} ${h / 2},${w} ${h / 2}' fill='none' stroke='${border.color}' stroke-width='${border.width}' stroke-dasharray='${border.strokeDasharray}' />`;
                    }
                    else if (shapType == "flowChartSummingJunction") {
                        var iDx, idy, il, ir, it, ib, hc = w / 2, vc = h / 2, wd2 = w / 2, hd2 = h / 2;
                        const angVal = Math.PI / 4;
                        iDx = wd2 * Math.cos(angVal);
                        idy = hd2 * Math.sin(angVal);
                        il = hc - iDx;
                        ir = hc + iDx;
                        it = vc - idy;
                        ib = vc + idy;
                        result += ` <polyline points='${il} ${it},${ir} ${ib}' fill='none' stroke='${border.color}' stroke-width='${border.width}' stroke-dasharray='${border.strokeDasharray}' />`;
                        result += ` <polyline points='${ir} ${it},${il} ${ib}' fill='none' stroke='${border.color}' stroke-width='${border.width}' stroke-dasharray='${border.strokeDasharray}' />`;
                    }
                    break;
                }
                case "roundRect":
                case "round1Rect":
                case "round2DiagRect":
                case "round2SameRect":
                case "snip1Rect":
                case "snip2DiagRect":
                case "snip2SameRect":
                case "flowChartAlternateProcess":
                case "flowChartPunchedCard": {
                    const shapAdjst_ary = PPTXXmlUtils.getTextByPathList(node, ["p:spPr", "a:prstGeom", "a:avLst", "a:gd"]);
                    let sAdj1, sAdj1_val;
                    let sAdj2, sAdj2_val;
                    let shpTyp, adjTyp;
                    if (shapAdjst_ary !== undefined && shapAdjst_ary.constructor === Array) {
                        for (const i of shapAdjst_ary.keys()) {
                            const sAdj_name = PPTXXmlUtils.getTextByPathList(shapAdjst_ary[i], ["attrs", "name"]);
                            if (sAdj_name == "adj1") {
                                sAdj1 = PPTXXmlUtils.getTextByPathList(shapAdjst_ary[i], ["attrs", "fmla"]);
                                sAdj1_val = parseInt(sAdj1.substr(4)) / 50000;
                            }
                            else if (sAdj_name == "adj2") {
                                sAdj2 = PPTXXmlUtils.getTextByPathList(shapAdjst_ary[i], ["attrs", "fmla"]);
                                sAdj2_val = parseInt(sAdj2.substr(4)) / 50000;
                            }
                        }
                    }
                    else if (shapAdjst_ary !== undefined && shapAdjst_ary.constructor !== Array) {
                        const sAdj = PPTXXmlUtils.getTextByPathList(shapAdjst_ary, ["attrs", "fmla"]);
                        sAdj1_val = parseInt(sAdj.substr(4)) / 50000;
                        sAdj2_val = 0;
                    }
                    let tranglRott = "";
                    switch (shapType) {
                        case "roundRect":
                        case "flowChartAlternateProcess": {
                            shpTyp = "round";
                            adjTyp = "cornrAll";
                            if (sAdj1_val === undefined)
                                sAdj1_val = 0.33334;
                            sAdj2_val = 0;
                            break;
                        }
                        case "round1Rect": {
                            shpTyp = "round";
                            adjTyp = "cornr1";
                            if (sAdj1_val === undefined)
                                sAdj1_val = 0.33334;
                            sAdj2_val = 0;
                            break;
                        }
                        case "round2DiagRect": {
                            shpTyp = "round";
                            adjTyp = "diag";
                            if (sAdj1_val === undefined)
                                sAdj1_val = 0.33334;
                            if (sAdj2_val === undefined)
                                sAdj2_val = 0;
                            break;
                        }
                        case "round2SameRect": {
                            shpTyp = "round";
                            adjTyp = "cornr2";
                            if (sAdj1_val === undefined)
                                sAdj1_val = 0.33334;
                            if (sAdj2_val === undefined)
                                sAdj2_val = 0;
                            break;
                        }
                        case "snip1Rect":
                        case "flowChartPunchedCard": {
                            shpTyp = "snip";
                            adjTyp = "cornr1";
                            if (sAdj1_val === undefined)
                                sAdj1_val = 0.33334;
                            sAdj2_val = 0;
                            if (shapType == "flowChartPunchedCard") {
                                tranglRott = `transform='translate(${w},0) scale(-1,1)'`;
                            }
                            break;
                        }
                        case "snip2DiagRect": {
                            shpTyp = "snip";
                            adjTyp = "diag";
                            if (sAdj1_val === undefined)
                                sAdj1_val = 0;
                            if (sAdj2_val === undefined)
                                sAdj2_val = 0.33334;
                            break;
                        }
                        case "snip2SameRect": {
                            shpTyp = "snip";
                            adjTyp = "cornr2";
                            if (sAdj1_val === undefined)
                                sAdj1_val = 0.33334;
                            if (sAdj2_val === undefined)
                                sAdj2_val = 0;
                            break;
                        }
                    }
                    let d_val = PPTXShapeUtils.shapeSnipRoundRectAlt(w, h, sAdj1_val, sAdj2_val, shpTyp, adjTyp);
                    result += `<path ${tranglRott}  d='${d_val}'  fill='${(!imgFillFlg ? (grndFillFlg ? "url(#linGrd_" + shpId + ")" : fillColor) : "url(#imgPtrn_" + shpId + ")")}' stroke='${border.color}' stroke-width='${border.width}' stroke-dasharray='${border.strokeDasharray}' />`;
                    break;
                }
                case "snipRoundRect": {
                    const shapAdjst_ary = PPTXXmlUtils.getTextByPathList(node, ["p:spPr", "a:prstGeom", "a:avLst", "a:gd"]);
                    let sAdj1, sAdj1_val = 0.33334;
                    let sAdj2, sAdj2_val = 0.33334;
                    if (shapAdjst_ary !== undefined) {
                        for (const i of shapAdjst_ary.keys()) {
                            const sAdj_name = PPTXXmlUtils.getTextByPathList(shapAdjst_ary[i], ["attrs", "name"]);
                            if (sAdj_name == "adj1") {
                                sAdj1 = PPTXXmlUtils.getTextByPathList(shapAdjst_ary[i], ["attrs", "fmla"]);
                                sAdj1_val = parseInt(sAdj1.substr(4)) / 50000;
                            }
                            else if (sAdj_name == "adj2") {
                                sAdj2 = PPTXXmlUtils.getTextByPathList(shapAdjst_ary[i], ["attrs", "fmla"]);
                                sAdj2_val = parseInt(sAdj2.substr(4)) / 50000;
                            }
                        }
                    }
                    const radius = Math.min(w, h) * sAdj1_val;
                    const snipSize = Math.min(w, h) * sAdj2_val;
                    let d_val = `M0,${(h - radius)} Q0,${h} ${radius},${h} L${w},${h} Q${w},${h} ${w},${(h - radius)} L${w},${snipSize} L${(w - snipSize)},0 L${snipSize},0 L0,${(h - snipSize)} z`;
                    result += `<path   d='${d_val}'  fill='${(!imgFillFlg ? (grndFillFlg ? "url(#linGrd_" + shpId + ")" : fillColor) : "url(#imgPtrn_" + shpId + ")")}' stroke='${border.color}' stroke-width='${border.width}' stroke-dasharray='${border.strokeDasharray}' />`;
                    break;
                }
                case "bentConnector2": {
                    var d = "";
                    const bendW = (drawW !== undefined) ? drawW : w;
                    const bendH = (drawH !== undefined) ? drawH : h;
                    d = `M ${bendW} 0 L ${bendW} ${bendH} L 0 ${bendH}`;
                    result += `<path d='${d}' stroke='${border.color}' stroke-width='${border.width}' stroke-dasharray='${border.strokeDasharray}' fill='none' `;
                    if (headEndNodeAttrs !== undefined && (headEndNodeAttrs["type"] === "triangle" || headEndNodeAttrs["type"] === "arrow")) {
                        result += `marker-start='url(#markerTriangle_${shpId})' `;
                    }
                    if (tailEndNodeAttrs !== undefined && (tailEndNodeAttrs["type"] === "triangle" || tailEndNodeAttrs["type"] === "arrow")) {
                        result += `marker-end='url(#markerTriangle_${shpId})' `;
                    }
                    result += "/>";
                    break;
                }
                case "rtTriangle": {
                    result += ` <polygon points='0 0,0 ${h},${w} ${h}' fill='${(!imgFillFlg ? (grndFillFlg ? "url(#linGrd_" + shpId + ")" : fillColor) : "url(#imgPtrn_" + shpId + ")")}' stroke='${border.color}' stroke-width='${border.width}' stroke-dasharray='${border.strokeDasharray}' />`;
                    break;
                }
                case "triangle":
                case "flowChartExtract":
                case "flowChartMerge": {
                    const shapAdjst = PPTXXmlUtils.getTextByPathList(node, ["p:spPr", "a:prstGeom", "a:avLst", "a:gd", "attrs", "fmla"]);
                    let shapAdjst_val = 0.5;
                    if (shapAdjst !== undefined) {
                        shapAdjst_val = parseInt(shapAdjst.substr(4)) * SLIDE_FACTOR$1;
                    }
                    let tranglRott = "";
                    if (shapType == "flowChartMerge") {
                        tranglRott = `transform='rotate(180 ${w / 2},${h / 2})'`;
                    }
                    result += ` <polygon ${tranglRott} points='${(w * shapAdjst_val)} 0,0 ${h},${w} ${h}' fill='${(!imgFillFlg ? (grndFillFlg ? "url(#linGrd_" + shpId + ")" : fillColor) : "url(#imgPtrn_" + shpId + ")")}' stroke='${border.color}' stroke-width='${border.width}' stroke-dasharray='${border.strokeDasharray}' />`;
                    break;
                }
                case "diamond":
                case "flowChartDecision":
                case "flowChartSort": {
                    result += ` <polygon points='${(w / 2)} 0,0 ${(h / 2)},${(w / 2)} ${h},${w} ${(h / 2)}' fill='${(!imgFillFlg ? (grndFillFlg ? "url(#linGrd_" + shpId + ")" : fillColor) : "url(#imgPtrn_" + shpId + ")")}' stroke='${border.color}' stroke-width='${border.width}' stroke-dasharray='${border.strokeDasharray}' />`;
                    if (shapType == "flowChartSort") {
                        result += ` <polyline points='0 ${h / 2},${w} ${h / 2}' fill='none' stroke='${border.color}' stroke-width='${border.width}' stroke-dasharray='${border.strokeDasharray}' />`;
                    }
                    break;
                }
                case "trapezoid":
                case "flowChartManualOperation":
                case "flowChartManualInput": {
                    const shapAdjst = PPTXXmlUtils.getTextByPathList(node, ["p:spPr", "a:prstGeom", "a:avLst", "a:gd", "attrs", "fmla"]);
                    let adjst_val = 0.2;
                    let max_adj_const = 0.7407;
                    if (shapAdjst !== undefined) {
                        const adjst = parseInt(shapAdjst.substr(4)) * SLIDE_FACTOR$1;
                        adjst_val = (adjst * 0.5) / max_adj_const;
                    }
                    var cnstVal = 0;
                    let tranglRott = "";
                    if (shapType == "flowChartManualOperation") {
                        tranglRott = `transform='rotate(180 ${w / 2},${h / 2})'`;
                    }
                    if (shapType == "flowChartManualInput") {
                        adjst_val = 0;
                        cnstVal = h / 5;
                    }
                    result += ` <polygon ${tranglRott} points='${(w * adjst_val)} ${cnstVal},0 ${h},${w} ${h},${(1 - adjst_val) * w} 0' fill='${(!imgFillFlg ? (grndFillFlg ? "url(#linGrd_" + shpId + ")" : fillColor) : "url(#imgPtrn_" + shpId + ")")}' stroke='${border.color}' stroke-width='${border.width}' stroke-dasharray='${border.strokeDasharray}' />`;
                    break;
                }
                case "parallelogram":
                case "flowChartInputOutput": {
                    const shapAdjst = PPTXXmlUtils.getTextByPathList(node, ["p:spPr", "a:prstGeom", "a:avLst", "a:gd", "attrs", "fmla"]);
                    let adjst_val = 0.25;
                    let max_adj_const;
                    if (w > h) {
                        max_adj_const = w / h;
                    }
                    else {
                        max_adj_const = h / w;
                    }
                    if (shapAdjst !== undefined) {
                        const adjst = parseInt(shapAdjst.substr(4)) / 100000;
                        adjst_val = adjst / max_adj_const;
                    }
                    result += ` <polygon points='${adjst_val * w} 0,0 ${h},${(1 - adjst_val) * w} ${h},${w} 0' fill='${(!imgFillFlg ? (grndFillFlg ? "url(#linGrd_" + shpId + ")" : fillColor) : "url(#imgPtrn_" + shpId + ")")}' stroke='${border.color}' stroke-width='${border.width}' stroke-dasharray='${border.strokeDasharray}' />`;
                    break;
                }
                case "pentagon": {
                    result += ` <polygon points='${(0.5 * w)} 0,0 ${(0.375 * h)},${(0.15 * w)} ${h},${0.85 * w} ${h},${w} ${0.375 * h}' fill='${(!imgFillFlg ? (grndFillFlg ? "url(#linGrd_" + shpId + ")" : fillColor) : "url(#imgPtrn_" + shpId + ")")}' stroke='${border.color}' stroke-width='${border.width}' stroke-dasharray='${border.strokeDasharray}' />`;
                    break;
                }
                case "hexagon":
                case "flowChartPreparation": {
                    const shapAdjst = PPTXXmlUtils.getTextByPathList(node, ["p:spPr", "a:prstGeom", "a:avLst", "a:gd", "attrs", "fmla"]);
                    let adj = 25000 * SLIDE_FACTOR$1;
                    const vf = 115470 * SLIDE_FACTOR$1;
                    const cnstVal1 = 50000 * SLIDE_FACTOR$1;
                    const cnstVal2 = 100000 * SLIDE_FACTOR$1;
                    const angVal1 = 60 * Math.PI / 180;
                    if (shapAdjst !== undefined) {
                        adj = parseInt(shapAdjst.substr(4)) * SLIDE_FACTOR$1;
                    }
                    var maxAdj, a, shd2, x1, x2, dy1, y1, y2, vc = h / 2, hd2 = h / 2;
                    const ss = Math.min(w, h);
                    maxAdj = cnstVal1 * w / ss;
                    a = (adj < 0) ? 0 : (adj > maxAdj) ? maxAdj : adj;
                    shd2 = hd2 * vf / cnstVal2;
                    x1 = ss * a / cnstVal2;
                    x2 = w - x1;
                    dy1 = shd2 * Math.sin(angVal1);
                    y1 = vc - dy1;
                    y2 = vc + dy1;
                    var d = `M${0},${vc} L${x1},${y1} L${x2},${y1} L${w},${vc} L${x2},${y2} L${x1},${y2} z`;
                    result += `<path   d='${d}'  fill='${(!imgFillFlg ? (grndFillFlg ? "url(#linGrd_" + shpId + ")" : fillColor) : "url(#imgPtrn_" + shpId + ")")}' stroke='${border.color}' stroke-width='${border.width}' stroke-dasharray='${border.strokeDasharray}' />`;
                    break;
                }
                case "heptagon": {
                    result += ` <polygon points='${(0.5 * w)} 0,${w / 8} ${h / 4},0 ${(5 / 8) * h},${w / 4} ${h},${(3 / 4) * w} ${h},${w} ${(5 / 8) * h},${(7 / 8) * w} ${h / 4}' fill='${(!imgFillFlg ? (grndFillFlg ? "url(#linGrd_" + shpId + ")" : fillColor) : "url(#imgPtrn_" + shpId + ")")}' stroke='${border.color}' stroke-width='${border.width}' stroke-dasharray='${border.strokeDasharray}' />`;
                    break;
                }
                case "octagon": {
                    const shapAdjst = PPTXXmlUtils.getTextByPathList(node, ["p:spPr", "a:prstGeom", "a:avLst", "a:gd", "attrs", "fmla"]);
                    let adj1 = 0.25;
                    if (shapAdjst !== undefined) {
                        adj1 = parseInt(shapAdjst.substr(4)) / 100000;
                    }
                    let adj2 = (1 - adj1);
                    result += ` <polygon points='${adj1 * w} 0,0 ${adj1 * h},0 ${adj2 * h},${adj1 * w} ${h},${adj2 * w} ${h},${w} ${adj2 * h},${w} ${adj1 * h},${adj2 * w} 0' fill='${(!imgFillFlg ? (grndFillFlg ? "url(#linGrd_" + shpId + ")" : fillColor) : "url(#imgPtrn_" + shpId + ")")}' stroke='${border.color}' stroke-width='${border.width}' stroke-dasharray='${border.strokeDasharray}' />`;
                    break;
                }
                case "decagon": {
                    result += ` <polygon points='${(3 / 8) * w} 0,${w / 8} ${h / 8},0 ${h / 2},${w / 8} ${(7 / 8) * h},${(3 / 8) * w} ${h},${(5 / 8) * w} ${h},${(7 / 8) * w} ${(7 / 8) * h},${w} ${h / 2},${(7 / 8) * w} ${h / 8},${(5 / 8) * w} 0' fill='${(!imgFillFlg ? (grndFillFlg ? "url(#linGrd_" + shpId + ")" : fillColor) : "url(#imgPtrn_" + shpId + ")")}' stroke='${border.color}' stroke-width='${border.width}' stroke-dasharray='${border.strokeDasharray}' />`;
                    break;
                }
                case "dodecagon": {
                    result += ` <polygon points='${(3 / 8) * w} 0,${w / 8} ${h / 8},0 ${(3 / 8) * h},0 ${(5 / 8) * h},${w / 8} ${(7 / 8) * h},${(3 / 8) * w} ${h},${(5 / 8) * w} ${h},${(7 / 8) * w} ${(7 / 8) * h},${w} ${(5 / 8) * h},${w} ${(3 / 8) * h},${(7 / 8) * w} ${h / 8},${(5 / 8) * w} 0' fill='${(!imgFillFlg ? (grndFillFlg ? "url(#linGrd_" + shpId + ")" : fillColor) : "url(#imgPtrn_" + shpId + ")")}' stroke='${border.color}' stroke-width='${border.width}' stroke-dasharray='${border.strokeDasharray}' />`;
                    break;
                }
                case "star4":
                case "star5":
                case "star6":
                case "star7":
                case "star8":
                case "star10":
                case "star12":
                case "star16":
                case "star24":
                case "star32": {
                    result += renderStar(shapType, drawW, drawH, imgFillFlg, grndFillFlg, fillColor, border, shpId, shapeArcAlt, node);
                    break;
                }
                case "pie":
                case "pieWedge":
                case "arc":
                case "chord": {
                    result += renderPieShape(shapType, drawW, drawH, imgFillFlg, grndFillFlg, fillColor, border, shpId, node, oShadowSvgUrlStr);
                    break;
                }
                case "frame": {
                    const shapAdjst = PPTXXmlUtils.getTextByPathList(node, ["p:spPr", "a:prstGeom", "a:avLst", "a:gd", "attrs", "fmla"]);
                    let adj1 = 12500 * SLIDE_FACTOR$1;
                    const cnstVal1 = 50000 * SLIDE_FACTOR$1;
                    const cnstVal2 = 100000 * SLIDE_FACTOR$1;
                    if (shapAdjst !== undefined) {
                        adj1 = parseInt(shapAdjst.substr(4)) * SLIDE_FACTOR$1;
                    }
                    var a1, x1, x4, y4;
                    if (adj1 < 0)
                        a1 = 0;
                    else if (adj1 > cnstVal1)
                        a1 = cnstVal1;
                    else
                        a1 = adj1;
                    x1 = Math.min(w, h) * a1 / cnstVal2;
                    x4 = w - x1;
                    y4 = h - x1;
                    var d = `M${0},${0} L${w},${0} L${w},${h} L${0},${h} zM${x1},${x1} L${x1},${y4} L${x4},${y4} L${x4},${x1} z`;
                    result += `<path   d='${d}'  fill='${(!imgFillFlg ? (grndFillFlg ? "url(#linGrd_" + shpId + ")" : fillColor) : "url(#imgPtrn_" + shpId + ")")}' stroke='${border.color}' stroke-width='${border.width}' stroke-dasharray='${border.strokeDasharray}' />`;
                    break;
                }
                case "donut": {
                    const shapAdjst = PPTXXmlUtils.getTextByPathList(node, ["p:spPr", "a:prstGeom", "a:avLst", "a:gd", "attrs", "fmla"]);
                    let adj = 25000 * SLIDE_FACTOR$1;
                    const cnstVal1 = 50000 * SLIDE_FACTOR$1;
                    const cnstVal2 = 100000 * SLIDE_FACTOR$1;
                    if (shapAdjst !== undefined) {
                        adj = parseInt(shapAdjst.substr(4)) * SLIDE_FACTOR$1;
                    }
                    let a, dr, iwd2, ihd2;
                    if (adj < 0)
                        a = 0;
                    else if (adj > cnstVal1)
                        a = cnstVal1;
                    else
                        a = adj;
                    dr = Math.min(w, h) * a / cnstVal2;
                    iwd2 = w / 2 - dr;
                    ihd2 = h / 2 - dr;
                    var d = `M${0},${h / 2}${PPTXShapeUtils.shapeArc(w / 2, h / 2, w / 2, h / 2, 180, 270, false).replace("M", "L")}${PPTXShapeUtils.shapeArc(w / 2, h / 2, w / 2, h / 2, 270, 360, false).replace("M", "L")}${PPTXShapeUtils.shapeArc(w / 2, h / 2, w / 2, h / 2, 0, 90, false).replace("M", "L")}${PPTXShapeUtils.shapeArc(w / 2, h / 2, w / 2, h / 2, 90, 180, false).replace("M", "L")} zM${dr},${h / 2}${PPTXShapeUtils.shapeArc(w / 2, h / 2, iwd2, ihd2, 180, 90, false).replace("M", "L")}${PPTXShapeUtils.shapeArc(w / 2, h / 2, iwd2, ihd2, 90, 0, false).replace("M", "L")}${PPTXShapeUtils.shapeArc(w / 2, h / 2, iwd2, ihd2, 0, -90, false).replace("M", "L")}${PPTXShapeUtils.shapeArc(w / 2, h / 2, iwd2, ihd2, 270, 180, false).replace("M", "L")} z`;
                    result += `<path   d='${d}'  fill='${(!imgFillFlg ? (grndFillFlg ? "url(#linGrd_" + shpId + ")" : fillColor) : "url(#imgPtrn_" + shpId + ")")}' stroke='${border.color}' stroke-width='${border.width}' stroke-dasharray='${border.strokeDasharray}' ${oShadowSvgUrlStr} />`;
                    break;
                }
                case "noSmoking": {
                    const shapAdjst = PPTXXmlUtils.getTextByPathList(node, ["p:spPr", "a:prstGeom", "a:avLst", "a:gd", "attrs", "fmla"]);
                    let adj = 18750 * SLIDE_FACTOR$1;
                    const cnstVal1 = 50000 * SLIDE_FACTOR$1;
                    const cnstVal2 = 100000 * SLIDE_FACTOR$1;
                    if (shapAdjst !== undefined) {
                        adj = parseInt(shapAdjst.substr(4)) * SLIDE_FACTOR$1;
                    }
                    let a, dr, iwd2, ihd2, ang, ct, st, m, n;
                    if (adj < 0)
                        a = 0;
                    else if (adj > cnstVal1)
                        a = cnstVal1;
                    else
                        a = adj;
                    dr = Math.min(w, h) * a / cnstVal2;
                    iwd2 = w / 2 - dr;
                    ihd2 = h / 2 - dr;
                    ang = Math.atan(h / w);
                    ct = ihd2 * Math.cos(ang);
                    st = iwd2 * Math.sin(ang);
                    m = Math.sqrt(ct * ct + st * st);
                    n = iwd2 * ihd2 / m;
                    const drd2 = dr / 2;
                    const dang = Math.atan(drd2 / n);
                    let dang2 = dang * 2;
                    let swAng = -Math.PI + dang2;
                    const stAng1 = ang - dang;
                    let stAng2 = stAng1 - Math.PI;
                    const stAng1deg = stAng1 * 180 / Math.PI;
                    const stAng2deg = stAng2 * 180 / Math.PI;
                    const swAng2deg = swAng * 180 / Math.PI;
                    let dx1 = n * Math.cos(stAng1);
                    let dy1 = n * Math.sin(stAng1);
                    var x1 = w / 2 + dx1;
                    var y1 = h / 2 + dy1;
                    var x2 = w / 2 - dx1;
                    var y2 = h / 2 - dy1;
                    var d = `M${0},${h / 2}${shapeArcAlt(w / 2, h / 2, w / 2, h / 2, 180, 270, false).replace("M", "L")}${shapeArcAlt(w / 2, h / 2, w / 2, h / 2, 270, 360, false).replace("M", "L")}${shapeArcAlt(w / 2, h / 2, w / 2, h / 2, 0, 90, false).replace("M", "L")}${shapeArcAlt(w / 2, h / 2, w / 2, h / 2, 90, 180, false).replace("M", "L")} zM${x1},${y1}${shapeArcAlt(w / 2, h / 2, iwd2, ihd2, stAng1deg, (stAng1deg + swAng2deg), false).replace("M", "L")} zM${x2},${y2}${shapeArcAlt(w / 2, h / 2, iwd2, ihd2, stAng2deg, (stAng2deg + swAng2deg), false).replace("M", "L")} z`;
                    result += `<path   d='${d}'  fill='${(!imgFillFlg ? (grndFillFlg ? "url(#linGrd_" + shpId + ")" : fillColor) : "url(#imgPtrn_" + shpId + ")")}' stroke='${border.color}' stroke-width='${border.width}' stroke-dasharray='${border.strokeDasharray}' />`;
                    break;
                }
                case "halfFrame": {
                    const shapAdjst_ary = PPTXXmlUtils.getTextByPathList(node, ["p:spPr", "a:prstGeom", "a:avLst", "a:gd"]);
                    let sAdj1, sAdj1_val = 3.5;
                    let sAdj2, sAdj2_val = 3.5;
                    const cnsVal = 100000 * SLIDE_FACTOR$1;
                    if (shapAdjst_ary !== undefined) {
                        for (const i of shapAdjst_ary.keys()) {
                            const sAdj_name = PPTXXmlUtils.getTextByPathList(shapAdjst_ary[i], ["attrs", "name"]);
                            if (sAdj_name == "adj1") {
                                sAdj1 = PPTXXmlUtils.getTextByPathList(shapAdjst_ary[i], ["attrs", "fmla"]);
                                sAdj1_val = parseInt(sAdj1.substr(4)) * SLIDE_FACTOR$1;
                            }
                            else if (sAdj_name == "adj2") {
                                sAdj2 = PPTXXmlUtils.getTextByPathList(shapAdjst_ary[i], ["attrs", "fmla"]);
                                sAdj2_val = parseInt(sAdj2.substr(4)) * SLIDE_FACTOR$1;
                            }
                        }
                    }
                    const minWH = Math.min(w, h);
                    let maxAdj2 = (cnsVal * w) / minWH;
                    let a1, a2;
                    if (sAdj2_val < 0)
                        a2 = 0;
                    else if (sAdj2_val > maxAdj2)
                        a2 = maxAdj2;
                    else
                        a2 = sAdj2_val;
                    var x1 = (minWH * a2) / cnsVal;
                    const g1 = h * x1 / w;
                    let g2 = h - g1;
                    let maxAdj1 = (cnsVal * g2) / minWH;
                    if (sAdj1_val < 0)
                        a1 = 0;
                    else if (sAdj1_val > maxAdj1)
                        a1 = maxAdj1;
                    else
                        a1 = sAdj1_val;
                    var y1 = minWH * a1 / cnsVal;
                    var dx2 = y1 * w / h;
                    var x2 = w - dx2;
                    let dy2 = x1 * h / w;
                    var y2 = h - dy2;
                    var d = `M0,0 L${w},${0} L${x2},${y1} L${x1},${y1} L${x1},${y2} L0,${h} z`;
                    result += `<path   d='${d}'  fill='${(!imgFillFlg ? (grndFillFlg ? "url(#linGrd_" + shpId + ")" : fillColor) : "url(#imgPtrn_" + shpId + ")")}' stroke='${border.color}' stroke-width='${border.width}' stroke-dasharray='${border.strokeDasharray}' />`;
                    break;
                }
                case "bracePair":
                case "bracketPair":
                case "leftBrace":
                case "leftBracket":
                case "rightBrace":
                case "rightBracket": {
                    result += renderBracket(shapType, drawW, drawH, imgFillFlg, grndFillFlg, fillColor, border, shpId, node);
                    break;
                }
                case "moon": {
                    const shapAdjst = PPTXXmlUtils.getTextByPathList(node, ["p:spPr", "a:prstGeom", "a:avLst", "a:gd", "attrs", "fmla"]);
                    let adj = 0.5;
                    if (shapAdjst !== undefined) {
                        adj = parseInt(shapAdjst.substr(4)) / 100000;
                    }
                    var hd2, cd2, cd4;
                    hd2 = h / 2;
                    cd2 = 180;
                    cd4 = 90;
                    let adj2 = (1 - adj) * w;
                    var d = `M${w},${h}${PPTXShapeUtils.shapeArc(w, hd2, w, hd2, cd4, (cd4 + cd2), false).replace("M", "L")}${PPTXShapeUtils.shapeArc(w, hd2, adj2, hd2, (cd4 + cd2), cd4, false).replace("M", "L")} z`;
                    result += `<path   d='${d}'  fill='${(!imgFillFlg ? (grndFillFlg ? "url(#linGrd_" + shpId + ")" : fillColor) : "url(#imgPtrn_" + shpId + ")")}' stroke='${border.color}' stroke-width='${border.width}' stroke-dasharray='${border.strokeDasharray}' />`;
                    break;
                }
                case "corner": {
                    const shapAdjst_ary = PPTXXmlUtils.getTextByPathList(node, ["p:spPr", "a:prstGeom", "a:avLst", "a:gd"]);
                    let sAdj1, sAdj1_val = 50000 * SLIDE_FACTOR$1;
                    let sAdj2, sAdj2_val = 50000 * SLIDE_FACTOR$1;
                    const cnsVal = 100000 * SLIDE_FACTOR$1;
                    if (shapAdjst_ary !== undefined) {
                        for (const i of shapAdjst_ary.keys()) {
                            const sAdj_name = PPTXXmlUtils.getTextByPathList(shapAdjst_ary[i], ["attrs", "name"]);
                            if (sAdj_name == "adj1") {
                                sAdj1 = PPTXXmlUtils.getTextByPathList(shapAdjst_ary[i], ["attrs", "fmla"]);
                                sAdj1_val = parseInt(sAdj1.substr(4)) * SLIDE_FACTOR$1;
                            }
                            else if (sAdj_name == "adj2") {
                                sAdj2 = PPTXXmlUtils.getTextByPathList(shapAdjst_ary[i], ["attrs", "fmla"]);
                                sAdj2_val = parseInt(sAdj2.substr(4)) * SLIDE_FACTOR$1;
                            }
                        }
                    }
                    const minWH = Math.min(w, h);
                    let maxAdj1 = cnsVal * h / minWH;
                    let maxAdj2 = cnsVal * w / minWH;
                    var a1, a2, x1, dy1, y1;
                    if (sAdj1_val < 0)
                        a1 = 0;
                    else if (sAdj1_val > maxAdj1)
                        a1 = maxAdj1;
                    else
                        a1 = sAdj1_val;
                    if (sAdj2_val < 0)
                        a2 = 0;
                    else if (sAdj2_val > maxAdj2)
                        a2 = maxAdj2;
                    else
                        a2 = sAdj2_val;
                    x1 = minWH * a2 / cnsVal;
                    dy1 = minWH * a1 / cnsVal;
                    y1 = h - dy1;
                    var d = `M0,0 L${x1},${0} L${x1},${y1} L${w},${y1} L${w},${h} L0,${h} z`;
                    result += `<path   d='${d}'  fill='${(!imgFillFlg ? (grndFillFlg ? "url(#linGrd_" + shpId + ")" : fillColor) : "url(#imgPtrn_" + shpId + ")")}' stroke='${border.color}' stroke-width='${border.width}' stroke-dasharray='${border.strokeDasharray}' />`;
                    break;
                }
                case "diagStripe": {
                    const shapAdjst = PPTXXmlUtils.getTextByPathList(node, ["p:spPr", "a:prstGeom", "a:avLst", "a:gd", "attrs", "fmla"]);
                    let sAdj1_val = 50000 * SLIDE_FACTOR$1;
                    const cnsVal = 100000 * SLIDE_FACTOR$1;
                    if (shapAdjst !== undefined) {
                        sAdj1_val = parseInt(shapAdjst.substr(4)) * SLIDE_FACTOR$1;
                    }
                    var a1, x2, y2;
                    if (sAdj1_val < 0)
                        a1 = 0;
                    else if (sAdj1_val > cnsVal)
                        a1 = cnsVal;
                    else
                        a1 = sAdj1_val;
                    x2 = w * a1 / cnsVal;
                    y2 = h * a1 / cnsVal;
                    var d = `M${0},${y2} L${x2},${0} L${w},${0} L${0},${h} z`;
                    result += `<path   d='${d}'  fill='${(!imgFillFlg ? (grndFillFlg ? "url(#linGrd_" + shpId + ")" : fillColor) : "url(#imgPtrn_" + shpId + ")")}' stroke='${border.color}' stroke-width='${border.width}' stroke-dasharray='${border.strokeDasharray}' />`;
                    break;
                }
                case "gear6":
                case "gear9": {
                    var gearNum = shapType.substr(4), d;
                    if (gearNum == "6") {
                        d = shapeGear(w, h / 3.5, parseInt(gearNum));
                    }
                    else {
                        d = shapeGear(w, h / 3.5, parseInt(gearNum));
                    }
                    result += `<path   d='${d}' transform='rotate(20,${(3 / 7) * h},${(3 / 7) * h})' fill='${(!imgFillFlg ? (grndFillFlg ? "url(#linGrd_" + shpId + ")" : fillColor) : "url(#imgPtrn_" + shpId + ")")}' stroke='${border.color}' stroke-width='${border.width}' stroke-dasharray='${border.strokeDasharray}' />`;
                    break;
                }
                case "bentConnector3": {
                    const shapAdjst = PPTXXmlUtils.getTextByPathList(node, ["p:spPr", "a:prstGeom", "a:avLst", "a:gd", "attrs", "fmla"]);
                    let shapAdjst_val = 0.5;
                    const connectorW = (drawW !== undefined) ? drawW : w;
                    const connectorH = (drawH !== undefined) ? drawH : h;
                    if (shapAdjst !== undefined) {
                        shapAdjst_val = parseInt(shapAdjst.substr(4)) / 100000;
                        result += ` <polyline points='0 0,${(shapAdjst_val) * connectorW} 0,${(shapAdjst_val) * connectorW} ${connectorH},${connectorW} ${connectorH}' fill='transparent'' stroke='${border.color}' stroke-width='${border.width}' stroke-dasharray='${border.strokeDasharray}' `;
                        if (headEndNodeAttrs !== undefined && (headEndNodeAttrs["type"] === "triangle" || headEndNodeAttrs["type"] === "arrow")) {
                            result += `marker-start='url(#markerTriangle_${shpId})' `;
                        }
                        if (tailEndNodeAttrs !== undefined && (tailEndNodeAttrs["type"] === "triangle" || tailEndNodeAttrs["type"] === "arrow")) {
                            result += `marker-end='url(#markerTriangle_${shpId})' `;
                        }
                        result += "/>";
                    }
                    break;
                }
                case "plus": {
                    const shapAdjst = PPTXXmlUtils.getTextByPathList(node, ["p:spPr", "a:prstGeom", "a:avLst", "a:gd", "attrs", "fmla"]);
                    let adj1 = 0.25;
                    if (shapAdjst !== undefined) {
                        adj1 = parseInt(shapAdjst.substr(4)) / 100000;
                    }
                    let adj2 = (1 - adj1);
                    result += ` <polygon points='${adj1 * w} 0,${adj1 * w} ${adj1 * h},0 ${adj1 * h},0 ${adj2 * h},${adj1 * w} ${adj2 * h},${adj1 * w} ${h},${adj2 * w} ${h},${adj2 * w} ${adj2 * h},${w} ${adj2 * h},${+w} ${adj1 * h},${adj2 * w} ${adj1 * h},${adj2 * w} 0' fill='${(!imgFillFlg ? (grndFillFlg ? "url(#linGrd_" + shpId + ")" : fillColor) : "url(#imgPtrn_" + shpId + ")")}' stroke='${border.color}' stroke-width='${border.width}' stroke-dasharray='${border.strokeDasharray}' />`;
                    break;
                }
                case "teardrop": {
                    const shapAdjst = PPTXXmlUtils.getTextByPathList(node, ["p:spPr", "a:prstGeom", "a:avLst", "a:gd", "attrs", "fmla"]);
                    let adj1 = 100000 * SLIDE_FACTOR$1;
                    const cnsVal1 = adj1;
                    const cnsVal2 = 200000 * SLIDE_FACTOR$1;
                    if (shapAdjst !== undefined) {
                        adj1 = parseInt(shapAdjst.substr(4)) * SLIDE_FACTOR$1;
                    }
                    var a1, r2, tw, th, sw, sh, dx1, dy1, x1, y1, x2, y2, rd45;
                    if (adj1 < 0)
                        a1 = 0;
                    else if (adj1 > cnsVal2)
                        a1 = cnsVal2;
                    else
                        a1 = adj1;
                    r2 = Math.sqrt(2);
                    tw = r2 * (w / 2);
                    th = r2 * (h / 2);
                    sw = (tw * a1) / cnsVal1;
                    sh = (th * a1) / cnsVal1;
                    rd45 = (45 * (Math.PI) / 180);
                    dx1 = sw * (Math.cos(rd45));
                    dy1 = sh * (Math.cos(rd45));
                    x1 = (w / 2) + dx1;
                    y1 = (h / 2) - dy1;
                    x2 = ((w / 2) + x1) / 2;
                    y2 = ((h / 2) + y1) / 2;
                    let d_val = `${PPTXShapeUtils.shapeArc(w / 2, h / 2, w / 2, h / 2, 180, 270, false)}Q ${x2},0 ${x1},${y1}Q ${w},${y2} ${w},${h / 2}${PPTXShapeUtils.shapeArc(w / 2, h / 2, w / 2, h / 2, 0, 90, false).replace("M", "L")}${PPTXShapeUtils.shapeArc(w / 2, h / 2, w / 2, h / 2, 90, 180, false).replace("M", "L")} z`;
                    result += `<path   d='${d_val}' fill='${(!imgFillFlg ? (grndFillFlg ? "url(#linGrd_" + shpId + ")" : fillColor) : "url(#imgPtrn_" + shpId + ")")}' stroke='${border.color}' stroke-width='${border.width}' stroke-dasharray='${border.strokeDasharray}' />`;
                    break;
                }
                case "plaque": {
                    const shapAdjst = PPTXXmlUtils.getTextByPathList(node, ["p:spPr", "a:prstGeom", "a:avLst", "a:gd", "attrs", "fmla"]);
                    let adjVal = 25000;
                    if (shapAdjst !== undefined) {
                        adjVal = parseInt(shapAdjst.substr(4));
                    }
                    if (adjVal < 0)
                        adjVal = 0;
                    else if (adjVal > 50000)
                        adjVal = 50000;
                    let r = (adjVal / 100000) * Math.min(w, h);
                    let d_val = `M${r},0A${r} ${r} 0 0 1 0,${r}L0,${(h - r)}A${r} ${r} 0 0 1 ${r},${h}L${(w - r)},${h}A${r} ${r} 0 0 1 ${w},${(h - r)}L${w},${r}A${r} ${r} 0 0 1 ${(w - r)},0 z`;
                    result += `<path   d='${d_val}' fill='${(!imgFillFlg ? (grndFillFlg ? "url(#linGrd_" + shpId + ")" : fillColor) : "url(#imgPtrn_" + shpId + ")")}' stroke='${border.color}' stroke-width='${border.width}' stroke-dasharray='${border.strokeDasharray}' />`;
                    break;
                }
                case "sun": {
                    const shapAdjst = PPTXXmlUtils.getTextByPathList(node, ["p:spPr", "a:prstGeom", "a:avLst", "a:gd", "attrs", "fmla"]);
                    const refr = SLIDE_FACTOR$1;
                    let adj1 = 25000 * refr;
                    const cnstVal1 = 12500 * refr;
                    const cnstVal2 = 46875 * refr;
                    if (shapAdjst !== undefined) {
                        adj1 = parseInt(shapAdjst.substr(4)) * refr;
                    }
                    let a1;
                    if (adj1 < cnstVal1)
                        a1 = cnstVal1;
                    else if (adj1 > cnstVal2)
                        a1 = cnstVal2;
                    else
                        a1 = adj1;
                    const cnstVa3 = 50000 * refr;
                    const cnstVa4 = 100000 * refr;
                    let g0 = cnstVa3 - a1, g1 = g0 * (30274 * refr) / (32768 * refr), g2 = g0 * (12540 * refr) / (32768 * refr), g5 = cnstVa3 - g1, g6 = cnstVa3 - g2, g10 = g5 * 3 / 4, g11 = g6 * 3 / 4, g12 = g10 + 3662 * refr, g13 = g11 + 36620 * refr, g14 = g11 + 12500 * refr, g15 = cnstVa4 - g10, g16 = cnstVa4 - g12, g17 = cnstVa4 - g13, g18 = cnstVa4 - g14, ox1 = w * (18436 * refr) / (21600 * refr), oy1 = h * (3163 * refr) / (21600 * refr), ox2 = w * (3163 * refr) / (21600 * refr), oy2 = h * (18436 * refr) / (21600 * refr), x10 = w * g10 / cnstVa4, x12 = w * g12 / cnstVa4, x13 = w * g13 / cnstVa4, x14 = w * g14 / cnstVa4, x15 = w * g15 / cnstVa4, x16 = w * g16 / cnstVa4, x17 = w * g17 / cnstVa4, x18 = w * g18 / cnstVa4, x19 = w * a1 / cnstVa4, wR = w * g0 / cnstVa4, hR = h * g0 / cnstVa4, y10 = h * g10 / cnstVa4, y12 = h * g12 / cnstVa4, y13 = h * g13 / cnstVa4, y14 = h * g14 / cnstVa4, y15 = h * g15 / cnstVa4, y16 = h * g16 / cnstVa4, y17 = h * g17 / cnstVa4, y18 = h * g18 / cnstVa4;
                    let d_val = `M${w},${h / 2} L${x15},${y18} L${x15},${y14}z M${ox1},${oy1} L${x16},${y17} L${x13},${y12}z M${w / 2},${0} L${x18},${y10} L${x14},${y10}z M${ox2},${oy1} L${x17},${y12} L${x12},${y17}z M${0},${h / 2} L${x10},${y14} L${x10},${y18}z M${ox2},${oy2} L${x12},${y13} L${x17},${y16}z M${w / 2},${h} L${x14},${y15} L${x18},${y15}z M${ox1},${oy2} L${x13},${y16} L${x16},${y13} z M${x19},${h / 2}${PPTXShapeUtils.shapeArc(w / 2, h / 2, wR, hR, 180, 540, false).replace("M", "L")} z`;
                    result += `<path   d='${d_val}' fill='${(!imgFillFlg ? (grndFillFlg ? "url(#linGrd_" + shpId + ")" : fillColor) : "url(#imgPtrn_" + shpId + ")")}' stroke='${border.color}' stroke-width='${border.width}' stroke-dasharray='${border.strokeDasharray}' />`;
                    break;
                }
                case "heart": {
                    var dx1, dx2, x1, x2, x3, x4, y1;
                    dx1 = w * 49 / 48;
                    dx2 = w * 10 / 48;
                    x1 = w / 2 - dx1;
                    x2 = w / 2 - dx2;
                    x3 = w / 2 + dx2;
                    x4 = w / 2 + dx1;
                    y1 = -h / 3;
                    let d_val = `M${w / 2},${h / 4}C${x3},${y1} ${x4},${h / 4} ${w / 2},${h}C${x1},${h / 4} ${x2},${y1} ${w / 2},${h / 4} z`;
                    result += `<path   d='${d_val}' fill='${(!imgFillFlg ? (grndFillFlg ? "url(#linGrd_" + shpId + ")" : fillColor) : "url(#imgPtrn_" + shpId + ")")}' stroke='${border.color}' stroke-width='${border.width}' stroke-dasharray='${border.strokeDasharray}' />`;
                    break;
                }
                case "lightningBolt": {
                    var x1 = w * 5022 / 21600, x2 = w * 11050 / 21600, x3 = w * 8472 / 21600, x4 = w * 8757 / 21600, x5 = w * 10012 / 21600, x6 = w * 14767 / 21600, x7 = w * 12222 / 21600, x8 = w * 12860 / 21600, x9 = w * 13917 / 21600, x10 = w * 7602 / 21600, x11 = w * 16577 / 21600, y1 = h * 3890 / 21600, y2 = h * 6080 / 21600, y3 = h * 6797 / 21600, y4 = h * 7437 / 21600, y5 = h * 12877 / 21600, y6 = h * 9705 / 21600, y7 = h * 12007 / 21600, y8 = h * 13987 / 21600, y9 = h * 8382 / 21600, y11 = h * 14915 / 21600;
                    let d_val = `M${x3},${0} L${x8},${y2} L${x2},${y3} L${x11},${y7} L${x6},${y5} L${w},${h} L${x5},${y11} L${x7},${y8} L${x1},${y6} L${x10},${y9} L${0},${y1} z`;
                    result += `<path d='${d_val}' fill='${(!imgFillFlg ? (grndFillFlg ? "url(#linGrd_" + shpId + ")" : fillColor) : "url(#imgPtrn_" + shpId + ")")}' stroke='${border.color}' stroke-width='${border.width}' stroke-dasharray='${border.strokeDasharray}' />`;
                    break;
                }
                case "cube": {
                    const shapAdjst = PPTXXmlUtils.getTextByPathList(node, ["p:spPr", "a:prstGeom", "a:avLst", "a:gd", "attrs", "fmla"]);
                    const refr = SLIDE_FACTOR$1;
                    let adj = 25000 * refr;
                    if (shapAdjst !== undefined) {
                        adj = parseInt(shapAdjst.substr(4)) * refr;
                    }
                    let d_val;
                    const cnstVal2 = 100000 * refr;
                    const ss = Math.min(w, h);
                    var a, y1, y4, x4;
                    a = (adj < 0) ? 0 : (adj > cnstVal2) ? cnstVal2 : adj;
                    y1 = ss * a / cnstVal2;
                    y4 = h - y1;
                    x4 = w - y1;
                    d_val = `M${0},${y1} L${y1},${0} L${w},${0} L${w},${y4} L${x4},${h} L${0},${h} zM${0},${y1} L${x4},${y1} M${x4},${y1} L${w},${0}M${x4},${y1} L${x4},${h}`;
                    result += `<path d='${d_val}' fill='${(!imgFillFlg ? (grndFillFlg ? "url(#linGrd_" + shpId + ")" : fillColor) : "url(#imgPtrn_" + shpId + ")")}' stroke='${border.color}' stroke-width='${border.width}' stroke-dasharray='${border.strokeDasharray}' />`;
                    break;
                }
                case "bevel": {
                    const shapAdjst = PPTXXmlUtils.getTextByPathList(node, ["p:spPr", "a:prstGeom", "a:avLst", "a:gd", "attrs", "fmla"]);
                    const refr = SLIDE_FACTOR$1;
                    let adj = 12500 * refr;
                    if (shapAdjst !== undefined) {
                        adj = parseInt(shapAdjst.substr(4)) * refr;
                    }
                    let d_val;
                    const cnstVal1 = 50000 * refr;
                    const cnstVal2 = 100000 * refr;
                    const ss = Math.min(w, h);
                    var a, x1, x2, y2;
                    a = (adj < 0) ? 0 : (adj > cnstVal1) ? cnstVal1 : adj;
                    x1 = ss * a / cnstVal2;
                    x2 = w - x1;
                    y2 = h - x1;
                    d_val = `M${0},${0} L${w},${0} L${w},${h} L${0},${h} z M${x1},${x1} L${x2},${x1} L${x2},${y2} L${x1},${y2} z M${0},${0} L${x1},${x1} M${0},${h} L${x1},${y2} M${w},${0} L${x2},${x1} M${w},${h} L${x2},${y2}`;
                    result += `<path d='${d_val}' fill='${(!imgFillFlg ? (grndFillFlg ? "url(#linGrd_" + shpId + ")" : fillColor) : "url(#imgPtrn_" + shpId + ")")}' stroke='${border.color}' stroke-width='${border.width}' stroke-dasharray='${border.strokeDasharray}' />`;
                    break;
                }
                case "foldedCorner": {
                    const shapAdjst = PPTXXmlUtils.getTextByPathList(node, ["p:spPr", "a:prstGeom", "a:avLst", "a:gd", "attrs", "fmla"]);
                    const refr = SLIDE_FACTOR$1;
                    let adj = 16667 * refr;
                    if (shapAdjst !== undefined) {
                        adj = parseInt(shapAdjst.substr(4)) * refr;
                    }
                    let d_val;
                    const cnstVal1 = 50000 * refr;
                    const cnstVal2 = 100000 * refr;
                    const ss = Math.min(w, h);
                    var a, dy2, dy1, x1, x2, y2, y1;
                    a = (adj < 0) ? 0 : (adj > cnstVal1) ? cnstVal1 : adj;
                    dy2 = ss * a / cnstVal2;
                    dy1 = dy2 / 5;
                    x1 = w - dy2;
                    x2 = x1 + dy1;
                    y2 = h - dy2;
                    y1 = y2 + dy1;
                    d_val = `M${x1},${h} L${x2},${y1} L${w},${y2} L${x1},${h} L${0},${h} L${0},${0} L${w},${0} L${w},${y2}`;
                    result += `<path d='${d_val}' fill='${(!imgFillFlg ? (grndFillFlg ? "url(#linGrd_" + shpId + ")" : fillColor) : "url(#imgPtrn_" + shpId + ")")}' stroke='${border.color}' stroke-width='${border.width}' stroke-dasharray='${border.strokeDasharray}' />`;
                    break;
                }
                case "cloud":
                case "cloudCallout": {
                    function fmt(num) {
                        return parseFloat(num.toFixed(2));
                    }
                    function ellipseArc(cx, cy, rx, ry, startAngle, sweepAngle) {
                        const endAngle = startAngle + sweepAngle;
                        const startX = cx + rx * Math.cos(startAngle * Math.PI / 180);
                        const startY = cy + ry * Math.sin(startAngle * Math.PI / 180);
                        const endX = cx + rx * Math.cos(endAngle * Math.PI / 180);
                        const endY = cy + ry * Math.sin(endAngle * Math.PI / 180);
                        const largeArc = Math.abs(sweepAngle) > 180 ? 1 : 0;
                        const sweep = sweepAngle > 0 ? 1 : 0;
                        return {
                            start: { x: fmt(startX), y: fmt(startY) },
                            end: { x: fmt(endX), y: fmt(endY) },
                            path: `A ${fmt(rx)} ${fmt(ry)} 0 ${largeArc} ${sweep} ${fmt(endX)} ${fmt(endY)}`
                        };
                    }
                    const x0 = fmt(w * 3900 / 43200);
                    const y0 = fmt(h * 14370 / 43200);
                    const rX1 = fmt(w * 6753 / 43200), rY1 = fmt(h * 9190 / 43200);
                    const rX2 = fmt(w * 5333 / 43200), rY2 = fmt(h * 7267 / 43200);
                    const rX3 = fmt(w * 4365 / 43200), rY3 = fmt(h * 5945 / 43200);
                    const rX4 = fmt(w * 4857 / 43200), rY4 = fmt(h * 6595 / 43200);
                    const rY5 = fmt(h * 7273 / 43200);
                    const rX6 = fmt(w * 6775 / 43200), rY6 = fmt(h * 9220 / 43200);
                    const rX7 = fmt(w * 5785 / 43200), rY7 = fmt(h * 7867 / 43200);
                    const rX8 = fmt(w * 6752 / 43200), rY8 = fmt(h * 9215 / 43200);
                    const rX9 = fmt(w * 7720 / 43200), rY9 = fmt(h * 10543 / 43200);
                    const rX10 = fmt(w * 4360 / 43200), rY10 = fmt(h * 5918 / 43200);
                    const rX11 = fmt(w * 4345 / 43200);
                    const sA1 = -11429249 / 60000, wA1 = 7426832 / 60000;
                    const sA2 = -8646143 / 60000, wA2 = 5396714 / 60000;
                    const sA3 = -8748475 / 60000, wA3 = 5983381 / 60000;
                    const sA4 = -7859164 / 60000, wA4 = 7034504 / 60000;
                    const sA5 = -4722533 / 60000, wA5 = 6541615 / 60000;
                    const sA6 = -2776035 / 60000, wA6 = 7816140 / 60000;
                    const sA7 = 37501 / 60000, wA7 = 6842000 / 60000;
                    const sA8 = 1347096 / 60000, wA8 = 6910353 / 60000;
                    const sA9 = 3974558 / 60000, wA9 = 4542661 / 60000;
                    const sA10 = -16496525 / 60000, wA10 = 8804134 / 60000;
                    const sA11 = -14809710 / 60000, wA11 = 9151131 / 60000;
                    const cX0 = fmt(x0 - rX1 * Math.cos(sA1 * Math.PI / 180));
                    const cY0 = fmt(y0 - rY1 * Math.sin(sA1 * Math.PI / 180));
                    const arc1 = ellipseArc(cX0, cY0, rX1, rY1, sA1, wA1);
                    const cX1 = fmt(arc1.end.x - rX2 * Math.cos(sA2 * Math.PI / 180));
                    const cY1 = fmt(arc1.end.y - rY2 * Math.sin(sA2 * Math.PI / 180));
                    const arc2 = ellipseArc(cX1, cY1, rX2, rY2, sA2, wA2);
                    const cX2 = fmt(arc2.end.x - rX3 * Math.cos(sA3 * Math.PI / 180));
                    const cY2 = fmt(arc2.end.y - rY3 * Math.sin(sA3 * Math.PI / 180));
                    const arc3 = ellipseArc(cX2, cY2, rX3, rY3, sA3, wA3);
                    const cX3 = fmt(arc3.end.x - rX4 * Math.cos(sA4 * Math.PI / 180));
                    const cY3 = fmt(arc3.end.y - rY4 * Math.sin(sA4 * Math.PI / 180));
                    const arc4 = ellipseArc(cX3, cY3, rX4, rY4, sA4, wA4);
                    const cX4 = fmt(arc4.end.x - rX2 * Math.cos(sA5 * Math.PI / 180));
                    const cY4 = fmt(arc4.end.y - rY5 * Math.sin(sA5 * Math.PI / 180));
                    const arc5 = ellipseArc(cX4, cY4, rX2, rY5, sA5, wA5);
                    const cX5 = fmt(arc5.end.x - rX6 * Math.cos(sA6 * Math.PI / 180));
                    const cY5 = fmt(arc5.end.y - rY6 * Math.sin(sA6 * Math.PI / 180));
                    const arc6 = ellipseArc(cX5, cY5, rX6, rY6, sA6, wA6);
                    const cX6 = fmt(arc6.end.x - rX7 * Math.cos(sA7 * Math.PI / 180));
                    const cY6 = fmt(arc6.end.y - rY7 * Math.sin(sA7 * Math.PI / 180));
                    const arc7 = ellipseArc(cX6, cY6, rX7, rY7, sA7, wA7);
                    const cX7 = fmt(arc7.end.x - rX8 * Math.cos(sA8 * Math.PI / 180));
                    const cY7 = fmt(arc7.end.y - rY8 * Math.sin(sA8 * Math.PI / 180));
                    const arc8 = ellipseArc(cX7, cY7, rX8, rY8, sA8, wA8);
                    const cX8 = fmt(arc8.end.x - rX9 * Math.cos(sA9 * Math.PI / 180));
                    const cY8 = fmt(arc8.end.y - rY9 * Math.sin(sA9 * Math.PI / 180));
                    const arc9 = ellipseArc(cX8, cY8, rX9, rY9, sA9, wA9);
                    const cX9 = fmt(arc9.end.x - rX10 * Math.cos(sA10 * Math.PI / 180));
                    const cY9 = fmt(arc9.end.y - rY10 * Math.sin(sA10 * Math.PI / 180));
                    const arc10 = ellipseArc(cX9, cY9, rX10, rY10, sA10, wA10);
                    const cX10 = fmt(arc10.end.x - rX11 * Math.cos(sA11 * Math.PI / 180));
                    const cY10 = fmt(arc10.end.y - rY3 * Math.sin(sA11 * Math.PI / 180));
                    const arc11 = ellipseArc(cX10, cY10, rX11, rY3, sA11, wA11);
                    let d1 = `M${x0},${y0} ${arc1.path} ${arc2.path} ${arc3.path} ${arc4.path} ${arc5.path} ${arc6.path} ${arc7.path} ${arc8.path} ${arc9.path} ${arc10.path} ${arc11.path} z`;
                    if (shapType == "cloudCallout") {
                        const shapAdjst_ary = PPTXXmlUtils.getTextByPathList(node, ["p:spPr", "a:prstGeom", "a:avLst", "a:gd"]);
                        const refr = SLIDE_FACTOR$1;
                        let sAdj1, adj1 = -20833 * refr;
                        let sAdj2, adj2 = 62500 * refr;
                        if (shapAdjst_ary !== undefined) {
                            for (const i of shapAdjst_ary.keys()) {
                                const sAdj_name = PPTXXmlUtils.getTextByPathList(shapAdjst_ary[i], ["attrs", "name"]);
                                if (sAdj_name == "adj1") {
                                    sAdj1 = PPTXXmlUtils.getTextByPathList(shapAdjst_ary[i], ["attrs", "fmla"]);
                                    adj1 = parseInt(sAdj1.substr(4)) * refr;
                                }
                                else if (sAdj_name == "adj2") {
                                    sAdj2 = PPTXXmlUtils.getTextByPathList(shapAdjst_ary[i], ["attrs", "fmla"]);
                                    adj2 = parseInt(sAdj2.substr(4)) * refr;
                                }
                            }
                        }
                        let d_val;
                        const cnstVal2 = 100000 * refr;
                        const ss = Math.min(w, h);
                        var wd2 = w / 2, hd2 = h / 2;
                        let dxPos, dyPos, xPos, yPos, ht, wt, g2, g3, g4, g5, g6, g7, g8, g9, g10, g11, g12, g13, g14, g15, g16, g17, g18, g19, g20, g21, g22, g23, g24, g25, g26, x23, x24, x25;
                        dxPos = w * adj1 / cnstVal2;
                        dyPos = h * adj2 / cnstVal2;
                        xPos = wd2 + dxPos;
                        yPos = hd2 + dyPos;
                        ht = hd2 * Math.cos(Math.atan(dyPos / dxPos));
                        wt = wd2 * Math.sin(Math.atan(dyPos / dxPos));
                        g2 = wd2 * Math.cos(Math.atan(wt / ht));
                        g3 = hd2 * Math.sin(Math.atan(wt / ht));
                        if (adj1 >= 0) {
                            g4 = wd2 + g2;
                            g5 = hd2 + g3;
                        }
                        else {
                            g4 = wd2 - g2;
                            g5 = hd2 - g3;
                        }
                        g6 = g4 - xPos;
                        g7 = g5 - yPos;
                        g8 = Math.sqrt(g6 * g6 + g7 * g7);
                        g9 = ss * 6600 / 21600;
                        g10 = g8 - g9;
                        g11 = g10 / 3;
                        g12 = ss * 1800 / 21600;
                        g13 = g11 + g12;
                        g14 = g13 * g6 / g8;
                        g15 = g13 * g7 / g8;
                        g16 = g14 + xPos;
                        g17 = g15 + yPos;
                        g18 = ss * 4800 / 21600;
                        g19 = g11 * 2;
                        g20 = g18 + g19;
                        g21 = g20 * g6 / g8;
                        g22 = g20 * g7 / g8;
                        g23 = g21 + xPos;
                        g24 = g22 + yPos;
                        g25 = ss * 1200 / 21600;
                        g26 = ss * 600 / 21600;
                        x23 = xPos + g26;
                        x24 = g16 + g25;
                        x25 = g23 + g12;
                        d_val =
                            `${PPTXShapeUtils.shapeArc(x23 - g26, yPos, g26, g26, 0, 360, false)} z M${x24},${g17}${PPTXShapeUtils.shapeArc(x24 - g25, g17, g25, g25, 0, 360, false).replace("M", "L")} z M${x25},${g24}${PPTXShapeUtils.shapeArc(x25 - g12, g24, g12, g12, 0, 360, false).replace("M", "L")} z`;
                        d1 += d_val;
                    }
                    result += `<path d='${d1}' fill='${(!imgFillFlg ? (grndFillFlg ? "url(#linGrd_" + shpId + ")" : fillColor) : "url(#imgPtrn_" + shpId + ")")}' stroke='${border.color}' stroke-width='${border.width}' stroke-dasharray='${border.strokeDasharray}' />`;
                    break;
                }
                case "smileyFace":
                case "verticalScroll":
                case "horizontalScroll": {
                    result += renderMiscShape(shapType, drawW, drawH, imgFillFlg, grndFillFlg, fillColor, border, shpId, node);
                    break;
                }
                case "wedgeEllipseCallout": {
                    const shapAdjst_ary = PPTXXmlUtils.getTextByPathList(node, ["p:spPr", "a:prstGeom", "a:avLst", "a:gd"]);
                    const refr = SLIDE_FACTOR$1;
                    let sAdj1, adj1 = -20833 * refr;
                    let sAdj2, adj2 = 62500 * refr;
                    if (shapAdjst_ary !== undefined) {
                        for (const i of shapAdjst_ary.keys()) {
                            const sAdj_name = PPTXXmlUtils.getTextByPathList(shapAdjst_ary[i], ["attrs", "name"]);
                            if (sAdj_name == "adj1") {
                                sAdj1 = PPTXXmlUtils.getTextByPathList(shapAdjst_ary[i], ["attrs", "fmla"]);
                                adj1 = parseInt(sAdj1.substr(4)) * refr;
                            }
                            else if (sAdj_name == "adj2") {
                                sAdj2 = PPTXXmlUtils.getTextByPathList(shapAdjst_ary[i], ["attrs", "fmla"]);
                                adj2 = parseInt(sAdj2.substr(4)) * refr;
                            }
                        }
                    }
                    let d_val;
                    const cnstVal1 = 100000 * SLIDE_FACTOR$1;
                    const angVal1 = 11 * Math.PI / 180;
                    var dxPos, dyPos, xPos, yPos, sdx, sdy, pang, stAng, enAng, dx1, dy1, x1, y1, dx2, dy2, x2, y2, swAng2, swAng, vc = h / 2, hc = w / 2;
                    dxPos = w * adj1 / cnstVal1;
                    dyPos = h * adj2 / cnstVal1;
                    xPos = hc + dxPos;
                    yPos = vc + dyPos;
                    sdx = dxPos * h;
                    sdy = dyPos * w;
                    pang = Math.atan(sdy / sdx);
                    stAng = pang + angVal1;
                    enAng = pang - angVal1;
                    dx1 = hc * Math.cos(stAng);
                    dy1 = vc * Math.sin(stAng);
                    dx2 = hc * Math.cos(enAng);
                    dy2 = vc * Math.sin(enAng);
                    if (dxPos >= 0) {
                        x1 = hc + dx1;
                        y1 = vc + dy1;
                        x2 = hc + dx2;
                        y2 = vc + dy2;
                    }
                    else {
                        x1 = hc - dx1;
                        y1 = vc - dy1;
                        x2 = hc - dx2;
                        y2 = vc - dy2;
                    }
                    d_val = `M${x1},${y1} L${xPos},${yPos} L${x2},${y2}${PPTXShapeUtils.shapeArcAlt(hc, vc, hc, vc, 0, 360, true)}`;
                    result += `<path d='${d_val}'${cloudTransformAttr} fill='${(!imgFillFlg ? (grndFillFlg ? "url(#linGrd_" + shpId + ")" : fillColor) : "url(#imgPtrn_" + shpId + ")")}' stroke='${border.color}' stroke-width='${border.width}' stroke-dasharray='${border.strokeDasharray}' />`;
                    break;
                }
                case "wedgeRectCallout": {
                    const shapAdjst_ary = PPTXXmlUtils.getTextByPathList(node, ["p:spPr", "a:prstGeom", "a:avLst", "a:gd"]);
                    const refr = SLIDE_FACTOR$1;
                    let sAdj1, adj1 = -20833 * refr;
                    let sAdj2, adj2 = 62500 * refr;
                    if (shapAdjst_ary !== undefined) {
                        for (const i of shapAdjst_ary.keys()) {
                            const sAdj_name = PPTXXmlUtils.getTextByPathList(shapAdjst_ary[i], ["attrs", "name"]);
                            if (sAdj_name == "adj1") {
                                sAdj1 = PPTXXmlUtils.getTextByPathList(shapAdjst_ary[i], ["attrs", "fmla"]);
                                adj1 = parseInt(sAdj1.substr(4)) * refr;
                            }
                            else if (sAdj_name == "adj2") {
                                sAdj2 = PPTXXmlUtils.getTextByPathList(shapAdjst_ary[i], ["attrs", "fmla"]);
                                adj2 = parseInt(sAdj2.substr(4)) * refr;
                            }
                        }
                    }
                    let d_val;
                    const cnstVal1 = 100000 * SLIDE_FACTOR$1;
                    var dxPos, dyPos, xPos, yPos, dx, dy, dq, ady, adq, dz, xg1, xg2, x1, x2, yg1, yg2, y1, y2, t1, xl, t2, xt, t3, xr, t4, xb, t5, yl, t6, yt, t7, yr, t8, yb, vc = h / 2, hc = w / 2;
                    dxPos = w * adj1 / cnstVal1;
                    dyPos = h * adj2 / cnstVal1;
                    xPos = hc + dxPos;
                    yPos = vc + dyPos;
                    dx = xPos - hc;
                    dy = yPos - vc;
                    dq = dxPos * h / w;
                    ady = Math.abs(dyPos);
                    adq = Math.abs(dq);
                    dz = ady - adq;
                    xg1 = (dxPos > 0) ? 7 : 2;
                    xg2 = (dxPos > 0) ? 10 : 5;
                    x1 = w * xg1 / 12;
                    x2 = w * xg2 / 12;
                    yg1 = (dyPos > 0) ? 7 : 2;
                    yg2 = (dyPos > 0) ? 10 : 5;
                    y1 = h * yg1 / 12;
                    y2 = h * yg2 / 12;
                    t1 = (dxPos > 0) ? 0 : xPos;
                    xl = (dz > 0) ? 0 : t1;
                    t2 = (dyPos > 0) ? x1 : xPos;
                    xt = (dz > 0) ? t2 : x1;
                    t3 = (dxPos > 0) ? xPos : w;
                    xr = (dz > 0) ? w : t3;
                    t4 = (dyPos > 0) ? xPos : x1;
                    xb = (dz > 0) ? t4 : x1;
                    t5 = (dxPos > 0) ? y1 : yPos;
                    yl = (dz > 0) ? y1 : t5;
                    t6 = (dyPos > 0) ? 0 : yPos;
                    yt = (dz > 0) ? t6 : 0;
                    t7 = (dxPos > 0) ? yPos : y1;
                    yr = (dz > 0) ? y1 : t7;
                    t8 = (dyPos > 0) ? yPos : h;
                    yb = (dz > 0) ? t8 : h;
                    d_val = `M${0},${0} L${x1},${0} L${xt},${yt} L${x2},${0} L${w},${0} L${w},${y1} L${xr},${yr} L${w},${y2} L${w},${h} L${x2},${h} L${xb},${yb} L${x1},${h} L${0},${h} L${0},${y2} L${xl},${yl} L${0},${y1} z`;
                    result += `<path d='${d_val}' fill='${(!imgFillFlg ? (grndFillFlg ? "url(#linGrd_" + shpId + ")" : fillColor) : "url(#imgPtrn_" + shpId + ")")}' stroke='${border.color}' stroke-width='${border.width}' stroke-dasharray='${border.strokeDasharray}' />`;
                    break;
                }
                case "wedgeRoundRectCallout": {
                    const shapAdjst_ary = PPTXXmlUtils.getTextByPathList(node, ["p:spPr", "a:prstGeom", "a:avLst", "a:gd"]);
                    const refr = SLIDE_FACTOR$1;
                    let sAdj1, adj1 = -20833 * refr;
                    let sAdj2, adj2 = 62500 * refr;
                    let sAdj3, adj3 = 16667 * refr;
                    if (shapAdjst_ary !== undefined) {
                        for (const i of shapAdjst_ary.keys()) {
                            const sAdj_name = PPTXXmlUtils.getTextByPathList(shapAdjst_ary[i], ["attrs", "name"]);
                            if (sAdj_name == "adj1") {
                                sAdj1 = PPTXXmlUtils.getTextByPathList(shapAdjst_ary[i], ["attrs", "fmla"]);
                                adj1 = parseInt(sAdj1.substr(4)) * refr;
                            }
                            else if (sAdj_name == "adj2") {
                                sAdj2 = PPTXXmlUtils.getTextByPathList(shapAdjst_ary[i], ["attrs", "fmla"]);
                                adj2 = parseInt(sAdj2.substr(4)) * refr;
                            }
                            else if (sAdj_name == "adj3") {
                                sAdj3 = PPTXXmlUtils.getTextByPathList(shapAdjst_ary[i], ["attrs", "fmla"]);
                                adj3 = parseInt(sAdj3.substr(4)) * refr;
                            }
                        }
                    }
                    let d_val;
                    const cnstVal1 = 100000 * SLIDE_FACTOR$1;
                    const ss = Math.min(w, h);
                    var dxPos, dyPos, xPos, yPos, dq, ady, adq, dz, xg1, xg2, x1, x2, yg1, yg2, y1, y2, t1, xl, t2, xt, t3, xr, t4, xb, t5, yl, t6, yt, t7, yr, t8, yb, u1, u2, v2, vc = h / 2, hc = w / 2;
                    dxPos = w * adj1 / cnstVal1;
                    dyPos = h * adj2 / cnstVal1;
                    xPos = hc + dxPos;
                    yPos = vc + dyPos;
                    dq = dxPos * h / w;
                    ady = Math.abs(dyPos);
                    adq = Math.abs(dq);
                    dz = ady - adq;
                    xg1 = (dxPos > 0) ? 7 : 2;
                    xg2 = (dxPos > 0) ? 10 : 5;
                    x1 = w * xg1 / 12;
                    x2 = w * xg2 / 12;
                    yg1 = (dyPos > 0) ? 7 : 2;
                    yg2 = (dyPos > 0) ? 10 : 5;
                    y1 = h * yg1 / 12;
                    y2 = h * yg2 / 12;
                    t1 = (dxPos > 0) ? 0 : xPos;
                    xl = (dz > 0) ? 0 : t1;
                    t2 = (dyPos > 0) ? x1 : xPos;
                    xt = (dz > 0) ? t2 : x1;
                    t3 = (dxPos > 0) ? xPos : w;
                    xr = (dz > 0) ? w : t3;
                    t4 = (dyPos > 0) ? xPos : x1;
                    xb = (dz > 0) ? t4 : x1;
                    t5 = (dxPos > 0) ? y1 : yPos;
                    yl = (dz > 0) ? y1 : t5;
                    t6 = (dyPos > 0) ? 0 : yPos;
                    yt = (dz > 0) ? t6 : 0;
                    t7 = (dxPos > 0) ? yPos : y1;
                    yr = (dz > 0) ? y1 : t7;
                    t8 = (dyPos > 0) ? yPos : h;
                    yb = (dz > 0) ? t8 : h;
                    u1 = ss * adj3 / cnstVal1;
                    u2 = w - u1;
                    v2 = h - u1;
                    d_val = `M${0},${u1}${PPTXShapeUtils.shapeArc(u1, u1, u1, u1, 180, 270, false).replace("M", "L")} L${x1},${0} L${xt},${yt} L${x2},${0} L${u2},${0}${PPTXShapeUtils.shapeArc(u2, u1, u1, u1, 270, 360, false).replace("M", "L")} L${w},${y1} L${xr},${yr} L${w},${y2} L${w},${v2}${PPTXShapeUtils.shapeArc(u2, v2, u1, u1, 0, 90, false).replace("M", "L")} L${x2},${h} L${xb},${yb} L${x1},${h} L${u1},${h}${PPTXShapeUtils.shapeArc(u1, v2, u1, u1, 90, 180, false).replace("M", "L")} L${0},${y2} L${xl},${yl} L${0},${y1} z`;
                    result += `<path d='${d_val}' fill='${(!imgFillFlg ? (grndFillFlg ? "url(#linGrd_" + shpId + ")" : fillColor) : "url(#imgPtrn_" + shpId + ")")}' stroke='${border.color}' stroke-width='${border.width}' stroke-dasharray='${border.strokeDasharray}' />`;
                    break;
                }
                case "accentBorderCallout1":
                case "accentBorderCallout2":
                case "accentBorderCallout3":
                case "borderCallout1":
                case "borderCallout2":
                case "borderCallout3":
                case "accentCallout1":
                case "accentCallout2":
                case "accentCallout3":
                case "callout1":
                case "callout2":
                case "callout3": {
                    const shapAdjst_ary = PPTXXmlUtils.getTextByPathList(node, ["p:spPr", "a:prstGeom", "a:avLst", "a:gd"]);
                    const refr = SLIDE_FACTOR$1;
                    let sAdj1, adj1 = 18750 * refr;
                    let sAdj2, adj2 = -8333 * refr;
                    let sAdj3, adj3 = 18750 * refr;
                    let sAdj4, adj4 = -16667 * refr;
                    let sAdj5, adj5 = 100000 * refr;
                    let sAdj6, adj6 = -16667 * refr;
                    let sAdj7, adj7 = 112963 * refr;
                    let sAdj8, adj8 = -8333 * refr;
                    if (shapAdjst_ary !== undefined) {
                        for (const i of shapAdjst_ary.keys()) {
                            const sAdj_name = PPTXXmlUtils.getTextByPathList(shapAdjst_ary[i], ["attrs", "name"]);
                            if (sAdj_name == "adj1") {
                                sAdj1 = PPTXXmlUtils.getTextByPathList(shapAdjst_ary[i], ["attrs", "fmla"]);
                                adj1 = parseInt(sAdj1.substr(4)) * refr;
                            }
                            else if (sAdj_name == "adj2") {
                                sAdj2 = PPTXXmlUtils.getTextByPathList(shapAdjst_ary[i], ["attrs", "fmla"]);
                                adj2 = parseInt(sAdj2.substr(4)) * refr;
                            }
                            else if (sAdj_name == "adj3") {
                                sAdj3 = PPTXXmlUtils.getTextByPathList(shapAdjst_ary[i], ["attrs", "fmla"]);
                                adj3 = parseInt(sAdj3.substr(4)) * refr;
                            }
                            else if (sAdj_name == "adj4") {
                                sAdj4 = PPTXXmlUtils.getTextByPathList(shapAdjst_ary[i], ["attrs", "fmla"]);
                                adj4 = parseInt(sAdj4.substr(4)) * refr;
                            }
                            else if (sAdj_name == "adj5") {
                                sAdj5 = PPTXXmlUtils.getTextByPathList(shapAdjst_ary[i], ["attrs", "fmla"]);
                                adj5 = parseInt(sAdj5.substr(4)) * refr;
                            }
                            else if (sAdj_name == "adj6") {
                                sAdj6 = PPTXXmlUtils.getTextByPathList(shapAdjst_ary[i], ["attrs", "fmla"]);
                                adj6 = parseInt(sAdj6.substr(4)) * refr;
                            }
                            else if (sAdj_name == "adj7") {
                                sAdj7 = PPTXXmlUtils.getTextByPathList(shapAdjst_ary[i], ["attrs", "fmla"]);
                                adj7 = parseInt(sAdj7.substr(4)) * refr;
                            }
                            else if (sAdj_name == "adj8") {
                                sAdj8 = PPTXXmlUtils.getTextByPathList(shapAdjst_ary[i], ["attrs", "fmla"]);
                                adj8 = parseInt(sAdj8.substr(4)) * refr;
                            }
                        }
                    }
                    let d_val;
                    const cnstVal1 = 100000 * refr;
                    switch (shapType) {
                        case "borderCallout1":
                        case "callout1":
                            if (shapAdjst_ary === undefined) {
                                adj1 = 18750 * refr;
                                adj2 = -8333 * refr;
                                adj3 = 112500 * refr;
                                adj4 = -38333 * refr;
                            }
                            var y1, x1, y2, x2;
                            y1 = h * adj1 / cnstVal1;
                            x1 = w * adj2 / cnstVal1;
                            y2 = h * adj3 / cnstVal1;
                            x2 = w * adj4 / cnstVal1;
                            d_val = `M${0},${0} L${w},${0} L${w},${h} L${0},${h} z M${x1},${y1} L${x2},${y2}`;
                            break;
                        case "borderCallout2":
                        case "callout2":
                            if (shapAdjst_ary === undefined) {
                                adj1 = 18750 * refr;
                                adj2 = -8333 * refr;
                                adj3 = 18750 * refr;
                                adj4 = -16667 * refr;
                                adj5 = 112500 * refr;
                                adj6 = -46667 * refr;
                            }
                            var y1, x1, y2, x2, y3, x3;
                            y1 = h * adj1 / cnstVal1;
                            x1 = w * adj2 / cnstVal1;
                            y2 = h * adj3 / cnstVal1;
                            x2 = w * adj4 / cnstVal1;
                            y3 = h * adj5 / cnstVal1;
                            x3 = w * adj6 / cnstVal1;
                            d_val = `M${0},${0} L${w},${0} L${w},${h} L${0},${h} z M${x1},${y1} L${x2},${y2} L${x3},${y3} L${x2},${y2}`;
                            break;
                        case "borderCallout3":
                        case "callout3":
                            if (shapAdjst_ary === undefined) {
                                adj1 = 18750 * refr;
                                adj2 = -8333 * refr;
                                adj3 = 18750 * refr;
                                adj4 = -16667 * refr;
                                adj5 = 100000 * refr;
                                adj6 = -16667 * refr;
                                adj7 = 112963 * refr;
                                adj8 = -8333 * refr;
                            }
                            var y1, x1, y2, x2, y3, x3, y4, x4;
                            y1 = h * adj1 / cnstVal1;
                            x1 = w * adj2 / cnstVal1;
                            y2 = h * adj3 / cnstVal1;
                            x2 = w * adj4 / cnstVal1;
                            y3 = h * adj5 / cnstVal1;
                            x3 = w * adj6 / cnstVal1;
                            y4 = h * adj7 / cnstVal1;
                            x4 = w * adj8 / cnstVal1;
                            d_val = `M${0},${0} L${w},${0} L${w},${h} L${0},${h} z M${x1},${y1} L${x2},${y2} L${x3},${y3} L${x4},${y4} L${x3},${y3} L${x2},${y2}`;
                            break;
                        case "accentBorderCallout1":
                        case "accentCallout1":
                            if (shapAdjst_ary === undefined) {
                                adj1 = 18750 * refr;
                                adj2 = -8333 * refr;
                                adj3 = 112500 * refr;
                                adj4 = -38333 * refr;
                            }
                            var y1, x1, y2, x2;
                            y1 = h * adj1 / cnstVal1;
                            x1 = w * adj2 / cnstVal1;
                            y2 = h * adj3 / cnstVal1;
                            x2 = w * adj4 / cnstVal1;
                            d_val = `M${0},${0} L${w},${0} L${w},${h} L${0},${h} z M${x1},${y1} L${x2},${y2} M${x1},${0} L${x1},${h}`;
                            break;
                        case "accentBorderCallout2":
                        case "accentCallout2":
                            if (shapAdjst_ary === undefined) {
                                adj1 = 18750 * refr;
                                adj2 = -8333 * refr;
                                adj3 = 18750 * refr;
                                adj4 = -16667 * refr;
                                adj5 = 112500 * refr;
                                adj6 = -46667 * refr;
                            }
                            var y1, x1, y2, x2, y3, x3;
                            y1 = h * adj1 / cnstVal1;
                            x1 = w * adj2 / cnstVal1;
                            y2 = h * adj3 / cnstVal1;
                            x2 = w * adj4 / cnstVal1;
                            y3 = h * adj5 / cnstVal1;
                            x3 = w * adj6 / cnstVal1;
                            d_val = `M${0},${0} L${w},${0} L${w},${h} L${0},${h} z M${x1},${y1} L${x2},${y2} L${x3},${y3} L${x2},${y2} M${x1},${0} L${x1},${h}`;
                            break;
                        case "accentBorderCallout3":
                        case "accentCallout3":
                            if (shapAdjst_ary === undefined) {
                                adj1 = 18750 * refr;
                                adj2 = -8333 * refr;
                                adj3 = 18750 * refr;
                                adj4 = -16667 * refr;
                                adj5 = 100000 * refr;
                                adj6 = -16667 * refr;
                                adj7 = 112963 * refr;
                                adj8 = -8333 * refr;
                            }
                            var y1, x1, y2, x2, y3, x3, y4, x4;
                            y1 = h * adj1 / cnstVal1;
                            x1 = w * adj2 / cnstVal1;
                            y2 = h * adj3 / cnstVal1;
                            x2 = w * adj4 / cnstVal1;
                            y3 = h * adj5 / cnstVal1;
                            x3 = w * adj6 / cnstVal1;
                            y4 = h * adj7 / cnstVal1;
                            x4 = w * adj8 / cnstVal1;
                            d_val = `M${0},${0} L${w},${0} L${w},${h} L${0},${h} z M${x1},${y1} L${x2},${y2} L${x3},${y3} L${x4},${y4} L${x3},${y3} L${x2},${y2} M${x1},${0} L${x1},${h}`;
                            break;
                    }
                    result += `<path d='${d_val}' fill='${(!imgFillFlg ? (grndFillFlg ? "url(#linGrd_" + shpId + ")" : fillColor) : "url(#imgPtrn_" + shpId + ")")}' stroke='${border.color}' stroke-width='${border.width}' stroke-dasharray='${border.strokeDasharray}' />`;
                    break;
                }
                case "leftRightRibbon": {
                    const shapAdjst_ary = PPTXXmlUtils.getTextByPathList(node, ["p:spPr", "a:prstGeom", "a:avLst", "a:gd"]);
                    const refr = SLIDE_FACTOR$1;
                    let sAdj1, adj1 = 50000 * refr;
                    let sAdj2, adj2 = 50000 * refr;
                    let sAdj3, adj3 = 16667 * refr;
                    if (shapAdjst_ary !== undefined) {
                        for (const i of shapAdjst_ary.keys()) {
                            const sAdj_name = PPTXXmlUtils.getTextByPathList(shapAdjst_ary[i], ["attrs", "name"]);
                            if (sAdj_name == "adj1") {
                                sAdj1 = PPTXXmlUtils.getTextByPathList(shapAdjst_ary[i], ["attrs", "fmla"]);
                                adj1 = parseInt(sAdj1.substr(4)) * refr;
                            }
                            else if (sAdj_name == "adj2") {
                                sAdj2 = PPTXXmlUtils.getTextByPathList(shapAdjst_ary[i], ["attrs", "fmla"]);
                                adj2 = parseInt(sAdj2.substr(4)) * refr;
                            }
                            else if (sAdj_name == "adj3") {
                                sAdj3 = PPTXXmlUtils.getTextByPathList(shapAdjst_ary[i], ["attrs", "fmla"]);
                                adj3 = parseInt(sAdj3.substr(4)) * refr;
                            }
                        }
                    }
                    let d_val;
                    const cnstVal1 = 33333 * refr;
                    const cnstVal2 = 100000 * refr;
                    const cnstVal3 = 200000 * refr;
                    const cnstVal4 = 400000 * refr;
                    const ss = Math.min(w, h);
                    var a3, maxAdj1, a1, w1, maxAdj2, a2, x1, x4, dy1, dy2, ly1, ry4, ly2, ry3, ly4, ry1, ly3, ry2, hR, x2, x3, y1, y2, wd32 = w / 32, vc = h / 2, hc = w / 2;
                    a3 = (adj3 < 0) ? 0 : (adj3 > cnstVal1) ? cnstVal1 : adj3;
                    maxAdj1 = cnstVal2 - a3;
                    a1 = (adj1 < 0) ? 0 : (adj1 > maxAdj1) ? maxAdj1 : adj1;
                    w1 = hc - wd32;
                    maxAdj2 = cnstVal2 * w1 / ss;
                    a2 = (adj2 < 0) ? 0 : (adj2 > maxAdj2) ? maxAdj2 : adj2;
                    x1 = ss * a2 / cnstVal2;
                    x4 = w - x1;
                    dy1 = h * a1 / cnstVal3;
                    dy2 = h * a3 / -cnstVal3;
                    ly1 = vc + dy2 - dy1;
                    ry4 = vc + dy1 - dy2;
                    ly2 = ly1 + dy1;
                    ry3 = h - ly2;
                    ly4 = ly2 * 2;
                    ry1 = h - ly4;
                    ly3 = ly4 - ly1;
                    ry2 = h - ly3;
                    hR = a3 * ss / cnstVal4;
                    x2 = hc - wd32;
                    x3 = hc + wd32;
                    y1 = ly1 + hR;
                    y2 = ry2 - hR;
                    d_val = `M${0},${ly2}L${x1},${0}L${x1},${ly1}L${hc},${ly1}${PPTXShapeUtils.shapeArcAlt(hc, y1, wd32, hR, 270, 450, false).replace("M", "L")}${PPTXShapeUtils.shapeArcAlt(hc, y2, wd32, hR, 270, 90, false).replace("M", "L")}L${x4},${ry2}L${x4},${ry1}L${w},${ry3}L${x4},${h}L${x4},${ry4}L${hc},${ry4}${PPTXShapeUtils.shapeArc(hc, ry4 - hR, wd32, hR, 90, 180, false).replace("M", "L")}L${x2},${ly3}L${x1},${ly3}L${x1},${ly4} zM${x3},${y1}L${x3},${ry2}M${x2},${y2}L${x2},${ly3}`;
                    result += `<path d='${d_val}' fill='${(!imgFillFlg ? (grndFillFlg ? "url(#linGrd_" + shpId + ")" : fillColor) : "url(#imgPtrn_" + shpId + ")")}' stroke='${border.color}' stroke-width='${border.width}' stroke-dasharray='${border.strokeDasharray}' />`;
                    break;
                }
                case "ribbon":
                case "ribbon2": {
                    const shapAdjst_ary = PPTXXmlUtils.getTextByPathList(node, ["p:spPr", "a:prstGeom", "a:avLst", "a:gd"]);
                    let sAdj1, adj1 = 16667 * SLIDE_FACTOR$1;
                    let sAdj2, adj2 = 50000 * SLIDE_FACTOR$1;
                    if (shapAdjst_ary !== undefined) {
                        for (const i of shapAdjst_ary.keys()) {
                            const sAdj_name = PPTXXmlUtils.getTextByPathList(shapAdjst_ary[i], ["attrs", "name"]);
                            if (sAdj_name == "adj1") {
                                sAdj1 = PPTXXmlUtils.getTextByPathList(shapAdjst_ary[i], ["attrs", "fmla"]);
                                adj1 = parseInt(sAdj1.substr(4)) * SLIDE_FACTOR$1;
                            }
                            else if (sAdj_name == "adj2") {
                                sAdj2 = PPTXXmlUtils.getTextByPathList(shapAdjst_ary[i], ["attrs", "fmla"]);
                                adj2 = parseInt(sAdj2.substr(4)) * SLIDE_FACTOR$1;
                            }
                        }
                    }
                    let d_val;
                    const cnstVal1 = 25000 * SLIDE_FACTOR$1;
                    const cnstVal2 = 33333 * SLIDE_FACTOR$1;
                    const cnstVal3 = 75000 * SLIDE_FACTOR$1;
                    const cnstVal4 = 100000 * SLIDE_FACTOR$1;
                    const cnstVal5 = 200000 * SLIDE_FACTOR$1;
                    const cnstVal6 = 400000 * SLIDE_FACTOR$1;
                    let hc = w / 2, t = 0, l = 0, b = h, r = w, wd8 = w / 8, wd32 = w / 32;
                    var a1, a2, x10, dx2, x2, x9, x3, x8, x5, x6, x4, x7, y1, y2, y4, y3, hR, y6;
                    a1 = (adj1 < 0) ? 0 : (adj1 > cnstVal2) ? cnstVal2 : adj1;
                    a2 = (adj2 < cnstVal1) ? cnstVal1 : (adj2 > cnstVal3) ? cnstVal3 : adj2;
                    x10 = r - wd8;
                    dx2 = w * a2 / cnstVal5;
                    x2 = hc - dx2;
                    x9 = hc + dx2;
                    x3 = x2 + wd32;
                    x8 = x9 - wd32;
                    x5 = x2 + wd8;
                    x6 = x9 - wd8;
                    x4 = x5 - wd32;
                    x7 = x6 + wd32;
                    hR = h * a1 / cnstVal6;
                    if (shapType == "ribbon2") {
                        let dy1, dy2, y7;
                        dy1 = h * a1 / cnstVal5;
                        y1 = b - dy1;
                        dy2 = h * a1 / cnstVal4;
                        y2 = b - dy2;
                        y4 = t + dy2;
                        y3 = (y4 + b) / 2;
                        y6 = b - hR;
                        y7 = y1 - hR;
                        d_val = `M${l},${b} L${wd8},${y3} L${l},${y4} L${x2},${y4} L${x2},${hR}${PPTXShapeUtils.shapeArcAlt(x3, hR, wd32, hR, 180, 270, false).replace("M", "L")} L${x8},${t}${PPTXShapeUtils.shapeArcAlt(x8, hR, wd32, hR, 270, 360, false).replace("M", "L")} L${x9},${y4} L${x9},${y4} L${r},${y4} L${x10},${y3} L${r},${b} L${x7},${b}${PPTXShapeUtils.shapeArc(x7, y6, wd32, hR, 90, 270, false).replace("M", "L")} L${x8},${y1}${PPTXShapeUtils.shapeArc(x8, y7, wd32, hR, 90, -90, false).replace("M", "L")} L${x3},${y2}${PPTXShapeUtils.shapeArc(x3, y7, wd32, hR, 270, 90, false).replace("M", "L")} L${x4},${y1}${PPTXShapeUtils.shapeArc(x4, y6, wd32, hR, 270, 450, false).replace("M", "L")} z M${x5},${y2} L${x5},${y6}M${x6},${y6} L${x6},${y2}M${x2},${y7} L${x2},${y4}M${x9},${y4} L${x9},${y7}`;
                    }
                    else if (shapType == "ribbon") {
                        let y5;
                        y1 = h * a1 / cnstVal5;
                        y2 = h * a1 / cnstVal4;
                        y4 = b - y2;
                        y3 = y4 / 2;
                        y5 = b - hR;
                        y6 = y2 - hR;
                        d_val = `M${l},${t} L${x4},${t}${PPTXShapeUtils.shapeArcAlt(x4, hR, wd32, hR, 270, 450, false).replace("M", "L")} L${x3},${y1}${PPTXShapeUtils.shapeArcAlt(x3, y6, wd32, hR, 270, 90, false).replace("M", "L")} L${x8},${y2}${PPTXShapeUtils.shapeArcAlt(x8, y6, wd32, hR, 90, -90, false).replace("M", "L")} L${x7},${y1}${PPTXShapeUtils.shapeArcAlt(x7, hR, wd32, hR, 90, 270, false).replace("M", "L")} L${r},${t} L${x10},${y3} L${r},${y4} L${x9},${y4} L${x9},${y5}${PPTXShapeUtils.shapeArc(x8, y5, wd32, hR, 0, 90, false).replace("M", "L")} L${x3},${b}${PPTXShapeUtils.shapeArc(x3, y5, wd32, hR, 90, 180, false).replace("M", "L")} L${x2},${y4} L${l},${y4} L${wd8},${y3} z M${x5},${hR} L${x5},${y2}M${x6},${y2} L${x6},${hR}M${x2},${y4} L${x2},${y6}M${x9},${y6} L${x9},${y4}`;
                    }
                    result += `<path d='${d_val}' fill='${(!imgFillFlg ? (grndFillFlg ? "url(#linGrd_" + shpId + ")" : fillColor) : "url(#imgPtrn_" + shpId + ")")}' stroke='${border.color}' stroke-width='${border.width}' stroke-dasharray='${border.strokeDasharray}' />`;
                    break;
                }
                case "doubleWave":
                case "wave": {
                    const shapAdjst_ary = PPTXXmlUtils.getTextByPathList(node, ["p:spPr", "a:prstGeom", "a:avLst", "a:gd"]);
                    let sAdj1, adj1 = (shapType == "doubleWave") ? 6250 * SLIDE_FACTOR$1 : 12500 * SLIDE_FACTOR$1;
                    let sAdj2, adj2 = 0;
                    if (shapAdjst_ary !== undefined) {
                        for (const i of shapAdjst_ary.keys()) {
                            const sAdj_name = PPTXXmlUtils.getTextByPathList(shapAdjst_ary[i], ["attrs", "name"]);
                            if (sAdj_name == "adj1") {
                                sAdj1 = PPTXXmlUtils.getTextByPathList(shapAdjst_ary[i], ["attrs", "fmla"]);
                                adj1 = parseInt(sAdj1.substr(4)) * SLIDE_FACTOR$1;
                            }
                            else if (sAdj_name == "adj2") {
                                sAdj2 = PPTXXmlUtils.getTextByPathList(shapAdjst_ary[i], ["attrs", "fmla"]);
                                adj2 = parseInt(sAdj2.substr(4)) * SLIDE_FACTOR$1;
                            }
                        }
                    }
                    let d_val;
                    const cnstVal2 = -1e4 * SLIDE_FACTOR$1;
                    const cnstVal3 = 50000 * SLIDE_FACTOR$1;
                    const cnstVal4 = 100000 * SLIDE_FACTOR$1;
                    let l = 0, b = h, r = w;
                    if (shapType == "doubleWave") {
                        const cnstVal1 = 12500 * SLIDE_FACTOR$1;
                        var a1, a2, y1, dy2, y2, y3, y4, y5, y6, of2, dx2, x2, dx8, x8, dx3, x3, dx4, x4, x5, x6, x7, x9, x15, x10, x11, x12, x13, x14;
                        a1 = (adj1 < 0) ? 0 : (adj1 > cnstVal1) ? cnstVal1 : adj1;
                        a2 = (adj2 < cnstVal2) ? cnstVal2 : (adj2 > cnstVal4) ? cnstVal4 : adj2;
                        y1 = h * a1 / cnstVal4;
                        dy2 = y1 * 10 / 3;
                        y2 = y1 - dy2;
                        y3 = y1 + dy2;
                        y4 = b - y1;
                        y5 = y4 - dy2;
                        y6 = y4 + dy2;
                        of2 = w * a2 / cnstVal3;
                        dx2 = (of2 > 0) ? 0 : of2;
                        x2 = l - dx2;
                        dx8 = (of2 > 0) ? of2 : 0;
                        x8 = r - dx8;
                        dx3 = (dx2 + x8) / 6;
                        x3 = x2 + dx3;
                        dx4 = (dx2 + x8) / 3;
                        x4 = x2 + dx4;
                        x5 = (x2 + x8) / 2;
                        x6 = x5 + dx3;
                        x7 = (x6 + x8) / 2;
                        x9 = l + dx8;
                        x15 = r + dx2;
                        x10 = x9 + dx3;
                        x11 = x9 + dx4;
                        x12 = (x9 + x15) / 2;
                        x13 = x12 + dx3;
                        x14 = (x13 + x15) / 2;
                        d_val = `M${x2},${y1} C${x3},${y2} ${x4},${y3} ${x5},${y1} C${x6},${y2} ${x7},${y3} ${x8},${y1} L${x15},${y4} C${x14},${y6} ${x13},${y5} ${x12},${y4} C${x11},${y6} ${x10},${y5} ${x9},${y4} z`;
                    }
                    else if (shapType == "wave") {
                        const cnstVal5 = 20000 * SLIDE_FACTOR$1;
                        var a1, a2, y1, dy2, y2, y3, y4, y5, y6, of2, dx2, x2, dx5, x5, dx3, x3, x4, x6, x10, x7, x8;
                        a1 = (adj1 < 0) ? 0 : (adj1 > cnstVal5) ? cnstVal5 : adj1;
                        a2 = (adj2 < cnstVal2) ? cnstVal2 : (adj2 > cnstVal4) ? cnstVal4 : adj2;
                        y1 = h * a1 / cnstVal4;
                        dy2 = y1 * 10 / 3;
                        y2 = y1 - dy2;
                        y3 = y1 + dy2;
                        y4 = b - y1;
                        y5 = y4 - dy2;
                        y6 = y4 + dy2;
                        of2 = w * a2 / cnstVal3;
                        dx2 = (of2 > 0) ? 0 : of2;
                        x2 = l - dx2;
                        dx5 = (of2 > 0) ? of2 : 0;
                        x5 = r - dx5;
                        dx3 = (dx2 + x5) / 3;
                        x3 = x2 + dx3;
                        x4 = (x3 + x5) / 2;
                        x6 = l + dx5;
                        x10 = r + dx2;
                        x7 = x6 + dx3;
                        x8 = (x7 + x10) / 2;
                        d_val = `M${x2},${y1} C${x3},${y2} ${x4},${y3} ${x5},${y1} L${x10},${y4} C${x8},${y6} ${x7},${y5} ${x6},${y4} z`;
                    }
                    result += `<path d='${d_val}' fill='${(!imgFillFlg ? (grndFillFlg ? "url(#linGrd_" + shpId + ")" : fillColor) : "url(#imgPtrn_" + shpId + ")")}' stroke='${border.color}' stroke-width='${border.width}' stroke-dasharray='${border.strokeDasharray}' />`;
                    break;
                }
                case "ellipseRibbon":
                case "ellipseRibbon2": {
                    const shapAdjst_ary = PPTXXmlUtils.getTextByPathList(node, ["p:spPr", "a:prstGeom", "a:avLst", "a:gd"]);
                    let sAdj1, adj1 = 25000 * SLIDE_FACTOR$1;
                    let sAdj2, adj2 = 50000 * SLIDE_FACTOR$1;
                    let sAdj3, adj3 = 12500 * SLIDE_FACTOR$1;
                    if (shapAdjst_ary !== undefined) {
                        for (const i of shapAdjst_ary.keys()) {
                            const sAdj_name = PPTXXmlUtils.getTextByPathList(shapAdjst_ary[i], ["attrs", "name"]);
                            if (sAdj_name == "adj1") {
                                sAdj1 = PPTXXmlUtils.getTextByPathList(shapAdjst_ary[i], ["attrs", "fmla"]);
                                adj1 = parseInt(sAdj1.substr(4)) * SLIDE_FACTOR$1;
                            }
                            else if (sAdj_name == "adj2") {
                                sAdj2 = PPTXXmlUtils.getTextByPathList(shapAdjst_ary[i], ["attrs", "fmla"]);
                                adj2 = parseInt(sAdj2.substr(4)) * SLIDE_FACTOR$1;
                            }
                            else if (sAdj_name == "adj3") {
                                sAdj3 = PPTXXmlUtils.getTextByPathList(shapAdjst_ary[i], ["attrs", "fmla"]);
                                adj3 = parseInt(sAdj3.substr(4)) * SLIDE_FACTOR$1;
                            }
                        }
                    }
                    let d_val;
                    const cnstVal1 = 25000 * SLIDE_FACTOR$1;
                    const cnstVal3 = 75000 * SLIDE_FACTOR$1;
                    const cnstVal4 = 100000 * SLIDE_FACTOR$1;
                    const cnstVal5 = 200000 * SLIDE_FACTOR$1;
                    let hc = w / 2, t = 0, l = 0, b = h, r = w, wd8 = w / 8;
                    var a1, a2, q10, q11, q12, minAdj3, a3, dx2, x2, x3, x4, x5, x6, dy1, f1, q1, q2, cx1, cx2, q1, dy3, q3, q4, q5, rh, q8, cx4, q9, cx5;
                    a1 = (adj1 < 0) ? 0 : (adj1 > cnstVal4) ? cnstVal4 : adj1;
                    a2 = (adj2 < cnstVal1) ? cnstVal1 : (adj2 > cnstVal3) ? cnstVal3 : adj2;
                    q10 = cnstVal4 - a1;
                    q11 = q10 / 2;
                    q12 = a1 - q11;
                    minAdj3 = (0 > q12) ? 0 : q12;
                    a3 = (adj3 < minAdj3) ? minAdj3 : (adj3 > a1) ? a1 : adj3;
                    dx2 = w * a2 / cnstVal5;
                    x2 = hc - dx2;
                    x3 = x2 + wd8;
                    x4 = r - x3;
                    x5 = r - x2;
                    x6 = r - wd8;
                    dy1 = h * a3 / cnstVal4;
                    f1 = 4 * dy1 / w;
                    q1 = x3 * x3 / w;
                    q2 = x3 - q1;
                    cx1 = x3 / 2;
                    cx2 = r - cx1;
                    q1 = h * a1 / cnstVal4;
                    dy3 = q1 - dy1;
                    q3 = x2 * x2 / w;
                    q4 = x2 - q3;
                    q5 = f1 * q4;
                    rh = b - q1;
                    q8 = dy1 * 14 / 16;
                    cx4 = x2 / 2;
                    q9 = f1 * cx4;
                    cx5 = r - cx4;
                    if (shapType == "ellipseRibbon") {
                        var y1, cy1, y3, q6, q7, cy3, y2, y5, y6, cy4, cy6, y7, y8;
                        y1 = f1 * q2;
                        cy1 = f1 * cx1;
                        y3 = q5 + dy3;
                        q6 = dy1 + dy3 - y3;
                        q7 = q6 + dy1;
                        cy3 = q7 + dy3;
                        y2 = (q8 + rh) / 2;
                        y5 = q5 + rh;
                        y6 = y3 + rh;
                        cy4 = q9 + rh;
                        cy6 = cy3 + rh;
                        y7 = y1 + dy3;
                        y8 = b - dy1;
                        d_val = `M${l},${t} Q${cx1},${cy1} ${x3},${y1} L${x2},${y3} Q${hc},${cy3} ${x5},${y3} L${x4},${y1} Q${cx2},${cy1} ${r},${t} L${x6},${y2} L${r},${rh} Q${cx5},${cy4} ${x5},${y5} L${x5},${y6} Q${hc},${cy6} ${x2},${y6} L${x2},${y5} Q${cx4},${cy4} ${l},${rh} L${wd8},${y2} zM${x2},${y5} L${x2},${y3}M${x5},${y3} L${x5},${y5}M${x3},${y1} L${x3},${y7}M${x4},${y7} L${x4},${y1}`;
                    }
                    else if (shapType == "ellipseRibbon2") {
                        var u1, y1, cu1, cy1, q3, q5, u3, y3, q6, q7, cu3, cy3, rh, q8, u2, y2, u5, y5, u6, y6, cu4, cy4, cu6, cy6, u7, y7;
                        u1 = f1 * q2;
                        y1 = b - u1;
                        cu1 = f1 * cx1;
                        cy1 = b - cu1;
                        u3 = q5 + dy3;
                        y3 = b - u3;
                        q6 = dy1 + dy3 - u3;
                        q7 = q6 + dy1;
                        cu3 = q7 + dy3;
                        cy3 = b - cu3;
                        u2 = (q8 + rh) / 2;
                        y2 = b - u2;
                        u5 = q5 + rh;
                        y5 = b - u5;
                        u6 = u3 + rh;
                        y6 = b - u6;
                        cu4 = q9 + rh;
                        cy4 = b - cu4;
                        cu6 = cu3 + rh;
                        cy6 = b - cu6;
                        u7 = u1 + dy3;
                        y7 = b - u7;
                        d_val = `M${l},${b} L${wd8},${y2} L${l},${q1} Q${cx4},${cy4} ${x2},${y5} L${x2},${y6} Q${hc},${cy6} ${x5},${y6} L${x5},${y5} Q${cx5},${cy4} ${r},${q1} L${x6},${y2} L${r},${b} Q${cx2},${cy1} ${x4},${y1} L${x5},${y3} Q${hc},${cy3} ${x2},${y3} L${x3},${y1} Q${cx1},${cy1} ${l},${b} zM${x2},${y3} L${x2},${y5}M${x5},${y5} L${x5},${y3}M${x3},${y7} L${x3},${y1}M${x4},${y1} L${x4},${y7}`;
                    }
                    result += `<path d='${d_val}' fill='${(!imgFillFlg ? (grndFillFlg ? "url(#linGrd_" + shpId + ")" : fillColor) : "url(#imgPtrn_" + shpId + ")")}' stroke='${border.color}' stroke-width='${border.width}' stroke-dasharray='${border.strokeDasharray}' />`;
                    break;
                }
                case "line":
                case "straightConnector1":
                case "bentConnector4":
                case "bentConnector5": {
                    let lineW = drawW;
                    let lineH = drawH;
                    if (lineW === undefined)
                        lineW = w;
                    if (lineH === undefined)
                        lineH = h;
                    var x1 = 0, y1 = 0, x2 = lineW, y2 = lineH;
                    result += `<line x1='${x1}' y1='${y1}' x2='${x2}' y2='${y2}' stroke='${border.color}' stroke-width='${border.width}' stroke-dasharray='${border.strokeDasharray}' `;
                    if (headEndNodeAttrs !== undefined && (headEndNodeAttrs["type"] === "triangle" || headEndNodeAttrs["type"] === "arrow")) {
                        result += `marker-start='url(#markerTriangle_${shpId})' `;
                    }
                    if (tailEndNodeAttrs !== undefined && (tailEndNodeAttrs["type"] === "triangle" || tailEndNodeAttrs["type"] === "arrow")) {
                        result += `marker-end='url(#markerTriangle_${shpId})' `;
                    }
                    result += "/>";
                    break;
                }
                case "curvedConnector2":
                case "curvedConnector3":
                case "curvedConnector4":
                case "curvedConnector5": {
                    const shapAdjst_ary = PPTXXmlUtils.getTextByPathList(node, ["p:spPr", "a:prstGeom", "a:avLst", "a:gd"]);
                    let adj1 = 50000;
                    if (shapAdjst_ary !== undefined) {
                        if (Array.isArray(shapAdjst_ary)) {
                            for (const i of shapAdjst_ary.keys()) {
                                const sAdj_name = PPTXXmlUtils.getTextByPathList(shapAdjst_ary[i], ["attrs", "name"]);
                                if (sAdj_name == "adj1") {
                                    let sAdj1 = PPTXXmlUtils.getTextByPathList(shapAdjst_ary[i], ["attrs", "fmla"]);
                                    adj1 = parseInt(sAdj1.substr(4));
                                    break;
                                }
                            }
                        }
                        else {
                            const sAdj_name = PPTXXmlUtils.getTextByPathList(shapAdjst_ary, ["attrs", "name"]);
                            if (sAdj_name == "adj1") {
                                let sAdj1 = PPTXXmlUtils.getTextByPathList(shapAdjst_ary, ["attrs", "fmla"]);
                                adj1 = parseInt(sAdj1.substr(4));
                            }
                        }
                    }
                    const curveW = (drawW !== undefined) ? drawW : w;
                    const curveH = (drawH !== undefined) ? drawH : h;
                    let cx1, cy1, cx2, cy2;
                    let pathD;
                    if (shapType === "curvedConnector2" || shapType === "curvedConnector3") {
                        const controlPointRatio = adj1 / 100000;
                        cx1 = curveW * controlPointRatio;
                        cy1 = 0;
                        cx2 = curveW * (1 - controlPointRatio);
                        cy2 = curveH;
                    }
                    else {
                        cx1 = curveW / 4;
                        cy1 = 0;
                        cx2 = curveW * 3 / 4;
                        cy2 = curveH;
                    }
                    pathD = `M 0,0 Q ${cx1},${cy1} ${curveW / 2},${curveH / 2} Q ${cx2},${cy2} ${curveW},${curveH}`;
                    result += `<path d='${pathD}' stroke='${border.color}' stroke-width='${border.width}' stroke-dasharray='${border.strokeDasharray}' fill='none' `;
                    if (headEndNodeAttrs !== undefined && (headEndNodeAttrs["type"] === "triangle" || headEndNodeAttrs["type"] === "arrow")) {
                        result += `marker-start='url(#markerTriangle_${shpId})' `;
                    }
                    if (tailEndNodeAttrs !== undefined && (tailEndNodeAttrs["type"] === "triangle" || tailEndNodeAttrs["type"] === "arrow")) {
                        result += `marker-end='url(#markerTriangle_${shpId})' `;
                    }
                    result += "/>";
                    break;
                }
                case "rightArrow":
                case "leftArrow":
                case "downArrow":
                case "upArrow":
                case "leftRightArrow":
                case "upDownArrow": {
                    result += renderArrow(shapType, drawW, drawH, imgFillFlg, grndFillFlg, fillColor, border, shpId, node);
                    break;
                }
                case "quadArrow": {
                    const shapAdjst_ary = PPTXXmlUtils.getTextByPathList(node, ["p:spPr", "a:prstGeom", "a:avLst", "a:gd"]);
                    let sAdj1, adj1 = 22500 * SLIDE_FACTOR$1;
                    let sAdj2, adj2 = 22500 * SLIDE_FACTOR$1;
                    let sAdj3, adj3 = 22500 * SLIDE_FACTOR$1;
                    const cnstVal1 = 50000 * SLIDE_FACTOR$1;
                    const cnstVal2 = 100000 * SLIDE_FACTOR$1;
                    const cnstVal3 = 200000 * SLIDE_FACTOR$1;
                    if (shapAdjst_ary !== undefined) {
                        for (const i of shapAdjst_ary.keys()) {
                            const sAdj_name = PPTXXmlUtils.getTextByPathList(shapAdjst_ary[i], ["attrs", "name"]);
                            if (sAdj_name == "adj1") {
                                sAdj1 = PPTXXmlUtils.getTextByPathList(shapAdjst_ary[i], ["attrs", "fmla"]);
                                adj1 = parseInt(sAdj1.substr(4)) * SLIDE_FACTOR$1;
                            }
                            else if (sAdj_name == "adj2") {
                                sAdj2 = PPTXXmlUtils.getTextByPathList(shapAdjst_ary[i], ["attrs", "fmla"]);
                                adj2 = parseInt(sAdj2.substr(4)) * SLIDE_FACTOR$1;
                            }
                            else if (sAdj_name == "adj3") {
                                sAdj3 = PPTXXmlUtils.getTextByPathList(shapAdjst_ary[i], ["attrs", "fmla"]);
                                adj3 = parseInt(sAdj3.substr(4)) * SLIDE_FACTOR$1;
                            }
                        }
                    }
                    var vc = h / 2, hc = w / 2, a1, a2, a3, q1, x1, x2, dx2, x3, dx3, x4, x5, x6, y2, y3, y4, y5, y6, maxAdj1, maxAdj3;
                    const minWH = Math.min(w, h);
                    if (adj2 < 0)
                        a2 = 0;
                    else if (adj2 > cnstVal1)
                        a2 = cnstVal1;
                    else
                        a2 = adj2;
                    maxAdj1 = 2 * a2;
                    if (adj1 < 0)
                        a1 = 0;
                    else if (adj1 > maxAdj1)
                        a1 = maxAdj1;
                    else
                        a1 = adj1;
                    q1 = cnstVal2 - maxAdj1;
                    maxAdj3 = q1 / 2;
                    if (adj3 < 0)
                        a3 = 0;
                    else if (adj3 > maxAdj3)
                        a3 = maxAdj3;
                    else
                        a3 = adj3;
                    x1 = minWH * a3 / cnstVal2;
                    dx2 = minWH * a2 / cnstVal2;
                    x2 = hc - dx2;
                    x5 = hc + dx2;
                    dx3 = minWH * a1 / cnstVal3;
                    x3 = hc - dx3;
                    x4 = hc + dx3;
                    x6 = w - x1;
                    y2 = vc - dx2;
                    y5 = vc + dx2;
                    y3 = vc - dx3;
                    y4 = vc + dx3;
                    y6 = h - x1;
                    let d_val = `M${0},${vc} L${x1},${y2} L${x1},${y3} L${x3},${y3} L${x3},${x1} L${x2},${x1} L${hc},${0} L${x5},${x1} L${x4},${x1} L${x4},${y3} L${x6},${y3} L${x6},${y2} L${w},${vc} L${x6},${y5} L${x6},${y4} L${x4},${y4} L${x4},${y6} L${x5},${y6} L${hc},${h} L${x2},${y6} L${x3},${y6} L${x3},${y4} L${x1},${y4} L${x1},${y5} z`;
                    result += `<path d='${d_val}' fill='${(!imgFillFlg ? (grndFillFlg ? "url(#linGrd_" + shpId + ")" : fillColor) : "url(#imgPtrn_" + shpId + ")")}' stroke='${border.color}' stroke-width='${border.width}' stroke-dasharray='${border.strokeDasharray}' />`;
                    break;
                }
                case "leftRightUpArrow": {
                    const shapAdjst_ary = PPTXXmlUtils.getTextByPathList(node, ["p:spPr", "a:prstGeom", "a:avLst", "a:gd"]);
                    let sAdj1, adj1 = 25000 * SLIDE_FACTOR$1;
                    let sAdj2, adj2 = 25000 * SLIDE_FACTOR$1;
                    let sAdj3, adj3 = 25000 * SLIDE_FACTOR$1;
                    const cnstVal1 = 50000 * SLIDE_FACTOR$1;
                    const cnstVal2 = 100000 * SLIDE_FACTOR$1;
                    const cnstVal3 = 200000 * SLIDE_FACTOR$1;
                    if (shapAdjst_ary !== undefined) {
                        for (const i of shapAdjst_ary.keys()) {
                            const sAdj_name = PPTXXmlUtils.getTextByPathList(shapAdjst_ary[i], ["attrs", "name"]);
                            if (sAdj_name == "adj1") {
                                sAdj1 = PPTXXmlUtils.getTextByPathList(shapAdjst_ary[i], ["attrs", "fmla"]);
                                adj1 = parseInt(sAdj1.substr(4)) * SLIDE_FACTOR$1;
                            }
                            else if (sAdj_name == "adj2") {
                                sAdj2 = PPTXXmlUtils.getTextByPathList(shapAdjst_ary[i], ["attrs", "fmla"]);
                                adj2 = parseInt(sAdj2.substr(4)) * SLIDE_FACTOR$1;
                            }
                            else if (sAdj_name == "adj3") {
                                sAdj3 = PPTXXmlUtils.getTextByPathList(shapAdjst_ary[i], ["attrs", "fmla"]);
                                adj3 = parseInt(sAdj3.substr(4)) * SLIDE_FACTOR$1;
                            }
                        }
                    }
                    var vc = h / 2, hc = w / 2, a1, a2, a3, q1, x1, x2, dx2, x3, dx3, x4, x5, x6, y2, dy2, y3, y4, y5, maxAdj1, maxAdj3;
                    const minWH = Math.min(w, h);
                    if (adj2 < 0)
                        a2 = 0;
                    else if (adj2 > cnstVal1)
                        a2 = cnstVal1;
                    else
                        a2 = adj2;
                    maxAdj1 = 2 * a2;
                    if (adj1 < 0)
                        a1 = 0;
                    else if (adj1 > maxAdj1)
                        a1 = maxAdj1;
                    else
                        a1 = adj1;
                    q1 = cnstVal2 - maxAdj1;
                    maxAdj3 = q1 / 2;
                    if (adj3 < 0)
                        a3 = 0;
                    else if (adj3 > maxAdj3)
                        a3 = maxAdj3;
                    else
                        a3 = adj3;
                    x1 = minWH * a3 / cnstVal2;
                    dx2 = minWH * a2 / cnstVal2;
                    x2 = hc - dx2;
                    x5 = hc + dx2;
                    dx3 = minWH * a1 / cnstVal3;
                    x3 = hc - dx3;
                    x4 = hc + dx3;
                    x6 = w - x1;
                    dy2 = minWH * a2 / cnstVal1;
                    y2 = h - dy2;
                    y4 = h - dx2;
                    y3 = y4 - dx3;
                    y5 = y4 + dx3;
                    let d_val = `M${0},${y4} L${x1},${y2} L${x1},${y3} L${x3},${y3} L${x3},${x1} L${x2},${x1} L${hc},${0} L${x5},${x1} L${x4},${x1} L${x4},${y3} L${x6},${y3} L${x6},${y2} L${w},${y4} L${x6},${h} L${x6},${y5} L${x1},${y5} L${x1},${h} z`;
                    result += `<path d='${d_val}' fill='${(!imgFillFlg ? (grndFillFlg ? "url(#linGrd_" + shpId + ")" : fillColor) : "url(#imgPtrn_" + shpId + ")")}' stroke='${border.color}' stroke-width='${border.width}' stroke-dasharray='${border.strokeDasharray}' />`;
                    break;
                }
                case "leftUpArrow": {
                    const shapAdjst_ary = PPTXXmlUtils.getTextByPathList(node, ["p:spPr", "a:prstGeom", "a:avLst", "a:gd"]);
                    let sAdj1, adj1 = 25000 * SLIDE_FACTOR$1;
                    let sAdj2, adj2 = 25000 * SLIDE_FACTOR$1;
                    let sAdj3, adj3 = 25000 * SLIDE_FACTOR$1;
                    const cnstVal1 = 50000 * SLIDE_FACTOR$1;
                    const cnstVal2 = 100000 * SLIDE_FACTOR$1;
                    const cnstVal3 = 200000 * SLIDE_FACTOR$1;
                    if (shapAdjst_ary !== undefined) {
                        for (const i of shapAdjst_ary.keys()) {
                            const sAdj_name = PPTXXmlUtils.getTextByPathList(shapAdjst_ary[i], ["attrs", "name"]);
                            if (sAdj_name == "adj1") {
                                sAdj1 = PPTXXmlUtils.getTextByPathList(shapAdjst_ary[i], ["attrs", "fmla"]);
                                adj1 = parseInt(sAdj1.substr(4)) * SLIDE_FACTOR$1;
                            }
                            else if (sAdj_name == "adj2") {
                                sAdj2 = PPTXXmlUtils.getTextByPathList(shapAdjst_ary[i], ["attrs", "fmla"]);
                                adj2 = parseInt(sAdj2.substr(4)) * SLIDE_FACTOR$1;
                            }
                            else if (sAdj_name == "adj3") {
                                sAdj3 = PPTXXmlUtils.getTextByPathList(shapAdjst_ary[i], ["attrs", "fmla"]);
                                adj3 = parseInt(sAdj3.substr(4)) * SLIDE_FACTOR$1;
                            }
                        }
                    }
                    var vc = h / 2, hc = w / 2, a1, a2, a3, x1, x2, dx4, dx3, x3, x4, x5, y2, y3, y4, y5, maxAdj1, maxAdj3;
                    const minWH = Math.min(w, h);
                    if (adj2 < 0)
                        a2 = 0;
                    else if (adj2 > cnstVal1)
                        a2 = cnstVal1;
                    else
                        a2 = adj2;
                    maxAdj1 = 2 * a2;
                    if (adj1 < 0)
                        a1 = 0;
                    else if (adj1 > maxAdj1)
                        a1 = maxAdj1;
                    else
                        a1 = adj1;
                    maxAdj3 = cnstVal2 - maxAdj1;
                    if (adj3 < 0)
                        a3 = 0;
                    else if (adj3 > maxAdj3)
                        a3 = maxAdj3;
                    else
                        a3 = adj3;
                    x1 = minWH * a3 / cnstVal2;
                    dx2 = minWH * a2 / cnstVal1;
                    x2 = w - dx2;
                    y2 = h - dx2;
                    dx4 = minWH * a2 / cnstVal2;
                    x4 = w - dx4;
                    y4 = h - dx4;
                    dx3 = minWH * a1 / cnstVal3;
                    x3 = x4 - dx3;
                    x5 = x4 + dx3;
                    y3 = y4 - dx3;
                    y5 = y4 + dx3;
                    let d_val = `M${0},${y4} L${x1},${y2} L${x1},${y3} L${x3},${y3} L${x3},${x1} L${x2},${x1} L${x4},${0} L${w},${x1} L${x5},${x1} L${x5},${y5} L${x1},${y5} L${x1},${h} z`;
                    result += `<path d='${d_val}' fill='${(!imgFillFlg ? (grndFillFlg ? "url(#linGrd_" + shpId + ")" : fillColor) : "url(#imgPtrn_" + shpId + ")")}' stroke='${border.color}' stroke-width='${border.width}' stroke-dasharray='${border.strokeDasharray}' />`;
                    break;
                }
                case "bentUpArrow": {
                    const shapAdjst_ary = PPTXXmlUtils.getTextByPathList(node, ["p:spPr", "a:prstGeom", "a:avLst", "a:gd"]);
                    let sAdj1, adj1 = 25000 * SLIDE_FACTOR$1;
                    let sAdj2, adj2 = 25000 * SLIDE_FACTOR$1;
                    let sAdj3, adj3 = 25000 * SLIDE_FACTOR$1;
                    const cnstVal1 = 50000 * SLIDE_FACTOR$1;
                    const cnstVal2 = 100000 * SLIDE_FACTOR$1;
                    const cnstVal3 = 200000 * SLIDE_FACTOR$1;
                    if (shapAdjst_ary !== undefined) {
                        for (const i of shapAdjst_ary.keys()) {
                            const sAdj_name = PPTXXmlUtils.getTextByPathList(shapAdjst_ary[i], ["attrs", "name"]);
                            if (sAdj_name == "adj1") {
                                sAdj1 = PPTXXmlUtils.getTextByPathList(shapAdjst_ary[i], ["attrs", "fmla"]);
                                adj1 = parseInt(sAdj1.substr(4)) * SLIDE_FACTOR$1;
                            }
                            else if (sAdj_name == "adj2") {
                                sAdj2 = PPTXXmlUtils.getTextByPathList(shapAdjst_ary[i], ["attrs", "fmla"]);
                                adj2 = parseInt(sAdj2.substr(4)) * SLIDE_FACTOR$1;
                            }
                            else if (sAdj_name == "adj3") {
                                sAdj3 = PPTXXmlUtils.getTextByPathList(shapAdjst_ary[i], ["attrs", "fmla"]);
                                adj3 = parseInt(sAdj3.substr(4)) * SLIDE_FACTOR$1;
                            }
                        }
                    }
                    var vc = h / 2, hc = w / 2, a1, a2, a3, dx1, x1, dx2, x2, dx3, x3, x4, y1, y2, dy2;
                    const minWH = Math.min(w, h);
                    if (adj1 < 0)
                        a1 = 0;
                    else if (adj1 > cnstVal1)
                        a1 = cnstVal1;
                    else
                        a1 = adj1;
                    if (adj2 < 0)
                        a2 = 0;
                    else if (adj2 > cnstVal1)
                        a2 = cnstVal1;
                    else
                        a2 = adj2;
                    if (adj3 < 0)
                        a3 = 0;
                    else if (adj3 > maxAdj3)
                        a3 = maxAdj3;
                    else
                        a3 = adj3;
                    y1 = minWH * a3 / cnstVal2;
                    dx1 = minWH * a2 / cnstVal1;
                    x1 = w - dx1;
                    dx3 = minWH * a2 / cnstVal2;
                    x3 = w - dx3;
                    dx2 = minWH * a1 / cnstVal3;
                    x2 = x3 - dx2;
                    x4 = x3 + dx2;
                    dy2 = minWH * a1 / cnstVal2;
                    y2 = h - dy2;
                    let d_val = `M${0},${y2} L${x2},${y2} L${x2},${y1} L${x1},${y1} L${x3},${0} L${w},${y1} L${x4},${y1} L${x4},${h} L${0},${h} z`;
                    result += `<path d='${d_val}' fill='${(!imgFillFlg ? (grndFillFlg ? "url(#linGrd_" + shpId + ")" : fillColor) : "url(#imgPtrn_" + shpId + ")")}' stroke='${border.color}' stroke-width='${border.width}' stroke-dasharray='${border.strokeDasharray}' />`;
                    break;
                }
                case "bentArrow": {
                    const shapAdjst_ary = PPTXXmlUtils.getTextByPathList(node, ["p:spPr", "a:prstGeom", "a:avLst", "a:gd"]);
                    let sAdj1, adj1 = 25000 * SLIDE_FACTOR$1;
                    let sAdj2, adj2 = 25000 * SLIDE_FACTOR$1;
                    let sAdj3, adj3 = 25000 * SLIDE_FACTOR$1;
                    let sAdj4, adj4 = 43750 * SLIDE_FACTOR$1;
                    const cnstVal1 = 50000 * SLIDE_FACTOR$1;
                    const cnstVal2 = 100000 * SLIDE_FACTOR$1;
                    if (shapAdjst_ary !== undefined) {
                        for (const i of shapAdjst_ary.keys()) {
                            const sAdj_name = PPTXXmlUtils.getTextByPathList(shapAdjst_ary[i], ["attrs", "name"]);
                            if (sAdj_name == "adj1") {
                                sAdj1 = PPTXXmlUtils.getTextByPathList(shapAdjst_ary[i], ["attrs", "fmla"]);
                                adj1 = parseInt(sAdj1.substr(4)) * SLIDE_FACTOR$1;
                            }
                            else if (sAdj_name == "adj2") {
                                sAdj2 = PPTXXmlUtils.getTextByPathList(shapAdjst_ary[i], ["attrs", "fmla"]);
                                adj2 = parseInt(sAdj2.substr(4)) * SLIDE_FACTOR$1;
                            }
                            else if (sAdj_name == "adj3") {
                                sAdj3 = PPTXXmlUtils.getTextByPathList(shapAdjst_ary[i], ["attrs", "fmla"]);
                                adj3 = parseInt(sAdj3.substr(4)) * SLIDE_FACTOR$1;
                            }
                            else if (sAdj_name == "adj4") {
                                sAdj4 = PPTXXmlUtils.getTextByPathList(shapAdjst_ary[i], ["attrs", "fmla"]);
                                adj4 = parseInt(sAdj4.substr(4)) * SLIDE_FACTOR$1;
                            }
                        }
                    }
                    var a1, a2, a3, a4, x3, x4, y3, y4, y5, y6, maxAdj1, maxAdj4;
                    const minWH = Math.min(w, h);
                    if (adj2 < 0)
                        a2 = 0;
                    else if (adj2 > cnstVal1)
                        a2 = cnstVal1;
                    else
                        a2 = adj2;
                    maxAdj1 = 2 * a2;
                    if (adj1 < 0)
                        a1 = 0;
                    else if (adj1 > maxAdj1)
                        a1 = maxAdj1;
                    else
                        a1 = adj1;
                    if (adj3 < 0)
                        a3 = 0;
                    else if (adj3 > cnstVal1)
                        a3 = cnstVal1;
                    else
                        a3 = adj3;
                    var th, aw2, th2, dh2, ah, bw, bh, bs, bd, bd3, bd2, th = minWH * a1 / cnstVal2;
                    aw2 = minWH * a2 / cnstVal2;
                    th2 = th / 2;
                    dh2 = aw2 - th2;
                    ah = minWH * a3 / cnstVal2;
                    bw = w - ah;
                    bh = h - dh2;
                    bs = (bw < bh) ? bw : bh;
                    maxAdj4 = cnstVal2 * bs / minWH;
                    if (adj4 < 0)
                        a4 = 0;
                    else if (adj4 > maxAdj4)
                        a4 = maxAdj4;
                    else
                        a4 = adj4;
                    bd = minWH * a4 / cnstVal2;
                    bd3 = bd - th;
                    bd2 = (bd3 > 0) ? bd3 : 0;
                    x3 = th + bd2;
                    x4 = w - ah;
                    y3 = dh2 + th;
                    y4 = y3 + dh2;
                    y5 = dh2 + bd;
                    y6 = y3 + bd2;
                    let d_val = `M${0},${h} L${0},${y5}${PPTXShapeUtils.shapeArc(bd, y5, bd, bd, 180, 270, false).replace("M", "L")} L${x4},${dh2} L${x4},${0} L${w},${aw2} L${x4},${y4} L${x4},${y3} L${x3},${y3}${PPTXShapeUtils.shapeArc(x3, y6, bd2, bd2, 270, 180, false).replace("M", "L")} L${th},${h} z`;
                    result += `<path d='${d_val}' fill='${(!imgFillFlg ? (grndFillFlg ? "url(#linGrd_" + shpId + ")" : fillColor) : "url(#imgPtrn_" + shpId + ")")}' stroke='${border.color}' stroke-width='${border.width}' stroke-dasharray='${border.strokeDasharray}' />`;
                    break;
                }
                case "uturnArrow": {
                    const shapAdjst_ary = PPTXXmlUtils.getTextByPathList(node, ["p:spPr", "a:prstGeom", "a:avLst", "a:gd"]);
                    let sAdj1, adj1 = 25000 * SLIDE_FACTOR$1;
                    let sAdj2, adj2 = 25000 * SLIDE_FACTOR$1;
                    let sAdj3, adj3 = 25000 * SLIDE_FACTOR$1;
                    let sAdj4, adj4 = 43750 * SLIDE_FACTOR$1;
                    let sAdj5, adj5 = 75000 * SLIDE_FACTOR$1;
                    const cnstVal1 = 25000 * SLIDE_FACTOR$1;
                    const cnstVal2 = 100000 * SLIDE_FACTOR$1;
                    if (shapAdjst_ary !== undefined) {
                        for (const i of shapAdjst_ary.keys()) {
                            const sAdj_name = PPTXXmlUtils.getTextByPathList(shapAdjst_ary[i], ["attrs", "name"]);
                            if (sAdj_name == "adj1") {
                                sAdj1 = PPTXXmlUtils.getTextByPathList(shapAdjst_ary[i], ["attrs", "fmla"]);
                                adj1 = parseInt(sAdj1.substr(4)) * SLIDE_FACTOR$1;
                            }
                            else if (sAdj_name == "adj2") {
                                sAdj2 = PPTXXmlUtils.getTextByPathList(shapAdjst_ary[i], ["attrs", "fmla"]);
                                adj2 = parseInt(sAdj2.substr(4)) * SLIDE_FACTOR$1;
                            }
                            else if (sAdj_name == "adj3") {
                                sAdj3 = PPTXXmlUtils.getTextByPathList(shapAdjst_ary[i], ["attrs", "fmla"]);
                                adj3 = parseInt(sAdj3.substr(4)) * SLIDE_FACTOR$1;
                            }
                            else if (sAdj_name == "adj4") {
                                sAdj4 = PPTXXmlUtils.getTextByPathList(shapAdjst_ary[i], ["attrs", "fmla"]);
                                adj4 = parseInt(sAdj4.substr(4)) * SLIDE_FACTOR$1;
                            }
                            else if (sAdj_name == "adj5") {
                                sAdj5 = PPTXXmlUtils.getTextByPathList(shapAdjst_ary[i], ["attrs", "fmla"]);
                                adj5 = parseInt(sAdj5.substr(4)) * SLIDE_FACTOR$1;
                            }
                        }
                    }
                    var a1, a2, a3, a4, a5, q1, q2, q3, x3, x4, x5, x6, x7, x8, x9, y4, y5, minAdj5, maxAdj1, maxAdj3, maxAdj4;
                    const minWH = Math.min(w, h);
                    if (adj2 < 0)
                        a2 = 0;
                    else if (adj2 > cnstVal1)
                        a2 = cnstVal1;
                    else
                        a2 = adj2;
                    maxAdj1 = 2 * a2;
                    if (adj1 < 0)
                        a1 = 0;
                    else if (adj1 > maxAdj1)
                        a1 = maxAdj1;
                    else
                        a1 = adj1;
                    q2 = a1 * minWH / h;
                    q3 = cnstVal2 - q2;
                    maxAdj3 = q3 * h / minWH;
                    if (adj3 < 0)
                        a3 = 0;
                    else if (adj3 > maxAdj3)
                        a3 = maxAdj3;
                    else
                        a3 = adj3;
                    q1 = a3 + a1;
                    minAdj5 = q1 * minWH / h;
                    if (adj5 < minAdj5)
                        a5 = minAdj5;
                    else if (adj5 > cnstVal2)
                        a5 = cnstVal2;
                    else
                        a5 = adj5;
                    var th, aw2, th2, dh2, ah, bw, bs, bd, bd3, bd2, th = minWH * a1 / cnstVal2;
                    aw2 = minWH * a2 / cnstVal2;
                    th2 = th / 2;
                    dh2 = aw2 - th2;
                    y5 = h * a5 / cnstVal2;
                    ah = minWH * a3 / cnstVal2;
                    y4 = y5 - ah;
                    x9 = w - dh2;
                    bw = x9 / 2;
                    bs = (bw < y4) ? bw : y4;
                    maxAdj4 = cnstVal2 * bs / minWH;
                    if (adj4 < 0)
                        a4 = 0;
                    else if (adj4 > maxAdj4)
                        a4 = maxAdj4;
                    else
                        a4 = adj4;
                    bd = minWH * a4 / cnstVal2;
                    bd3 = bd - th;
                    bd2 = (bd3 > 0) ? bd3 : 0;
                    x3 = th + bd2;
                    x8 = w - aw2;
                    x6 = x8 - aw2;
                    x7 = x6 + dh2;
                    x4 = x9 - bd;
                    x5 = x7 - bd2;
                    let d_val = `M${0},${h} L${0},${bd}${shapeArcAlt(bd, bd, bd, bd, 180, 270, false).replace("M", "L")} L${x4},${0}${shapeArcAlt(x4, bd, bd, bd, 270, 360, false).replace("M", "L")} L${x9},${y4} L${w},${y4} L${x8},${y5} L${x6},${y4} L${x7},${y4} L${x7},${x3}${shapeArcAlt(x5, x3, bd2, bd2, 0, -90, false).replace("M", "L")} L${x3},${th}${shapeArcAlt(x3, x3, bd2, bd2, 270, 180, false).replace("M", "L")} L${th},${h} z`;
                    result += `<path d='${d_val}' fill='${(!imgFillFlg ? (grndFillFlg ? "url(#linGrd_" + shpId + ")" : fillColor) : "url(#imgPtrn_" + shpId + ")")}' stroke='${border.color}' stroke-width='${border.width}' stroke-dasharray='${border.strokeDasharray}' />`;
                    break;
                }
                case "stripedRightArrow": {
                    const shapAdjst_ary = PPTXXmlUtils.getTextByPathList(node, ["p:spPr", "a:prstGeom", "a:avLst", "a:gd"]);
                    let sAdj1, adj1 = 50000 * SLIDE_FACTOR$1;
                    let sAdj2, adj2 = 50000 * SLIDE_FACTOR$1;
                    const cnstVal1 = 100000 * SLIDE_FACTOR$1;
                    const cnstVal2 = 200000 * SLIDE_FACTOR$1;
                    const cnstVal3 = 84375 * SLIDE_FACTOR$1;
                    if (shapAdjst_ary !== undefined) {
                        for (const i of shapAdjst_ary.keys()) {
                            const sAdj_name = PPTXXmlUtils.getTextByPathList(shapAdjst_ary[i], ["attrs", "name"]);
                            if (sAdj_name == "adj1") {
                                sAdj1 = PPTXXmlUtils.getTextByPathList(shapAdjst_ary[i], ["attrs", "fmla"]);
                                adj1 = parseInt(sAdj1.substr(4)) * SLIDE_FACTOR$1;
                            }
                            else if (sAdj_name == "adj2") {
                                sAdj2 = PPTXXmlUtils.getTextByPathList(shapAdjst_ary[i], ["attrs", "fmla"]);
                                adj2 = parseInt(sAdj2.substr(4)) * SLIDE_FACTOR$1;
                            }
                        }
                    }
                    var a1, a2, x4, x5, dx5, x6, y1, dy1, y2, maxAdj2, vc = h / 2;
                    const minWH = Math.min(w, h);
                    maxAdj2 = cnstVal3 * w / minWH;
                    if (adj1 < 0)
                        a1 = 0;
                    else if (adj1 > cnstVal1)
                        a1 = cnstVal1;
                    else
                        a1 = adj1;
                    if (adj2 < 0)
                        a2 = 0;
                    else if (adj2 > maxAdj2)
                        a2 = maxAdj2;
                    else
                        a2 = adj2;
                    x4 = minWH * 5 / 32;
                    dx5 = minWH * a2 / cnstVal1;
                    x5 = w - dx5;
                    dy1 = h * a1 / cnstVal2;
                    y1 = vc - dy1;
                    y2 = vc + dy1;
                    const ssd8 = minWH / 8, ssd16 = minWH / 16, ssd32 = minWH / 32;
                    let d_val = `M${0},${y1} L${ssd32},${y1} L${ssd32},${y2} L${0},${y2} z M${ssd16},${y1} L${ssd8},${y1} L${ssd8},${y2} L${ssd16},${y2} z M${x4},${y1} L${x5},${y1} L${x5},${0} L${w},${vc} L${x5},${h} L${x5},${y2} L${x4},${y2} z`;
                    result += `<path d='${d_val}' fill='${(!imgFillFlg ? (grndFillFlg ? "url(#linGrd_" + shpId + ")" : fillColor) : "url(#imgPtrn_" + shpId + ")")}' stroke='${border.color}' stroke-width='${border.width}' stroke-dasharray='${border.strokeDasharray}' />`;
                    break;
                }
                case "notchedRightArrow": {
                    const shapAdjst_ary = PPTXXmlUtils.getTextByPathList(node, ["p:spPr", "a:prstGeom", "a:avLst", "a:gd"]);
                    let sAdj1, adj1 = 50000 * SLIDE_FACTOR$1;
                    let sAdj2, adj2 = 50000 * SLIDE_FACTOR$1;
                    const cnstVal1 = 100000 * SLIDE_FACTOR$1;
                    const cnstVal2 = 200000 * SLIDE_FACTOR$1;
                    if (shapAdjst_ary !== undefined) {
                        for (const i of shapAdjst_ary.keys()) {
                            const sAdj_name = PPTXXmlUtils.getTextByPathList(shapAdjst_ary[i], ["attrs", "name"]);
                            if (sAdj_name == "adj1") {
                                sAdj1 = PPTXXmlUtils.getTextByPathList(shapAdjst_ary[i], ["attrs", "fmla"]);
                                adj1 = parseInt(sAdj1.substr(4)) * SLIDE_FACTOR$1;
                            }
                            else if (sAdj_name == "adj2") {
                                sAdj2 = PPTXXmlUtils.getTextByPathList(shapAdjst_ary[i], ["attrs", "fmla"]);
                                adj2 = parseInt(sAdj2.substr(4)) * SLIDE_FACTOR$1;
                            }
                        }
                    }
                    var a1, a2, x1, x2, dx2, y1, dy1, y2, maxAdj2, vc = h / 2, hd2 = vc;
                    const minWH = Math.min(w, h);
                    maxAdj2 = cnstVal1 * w / minWH;
                    if (adj1 < 0)
                        a1 = 0;
                    else if (adj1 > cnstVal1)
                        a1 = cnstVal1;
                    else
                        a1 = adj1;
                    if (adj2 < 0)
                        a2 = 0;
                    else if (adj2 > maxAdj2)
                        a2 = maxAdj2;
                    else
                        a2 = adj2;
                    dx2 = minWH * a2 / cnstVal1;
                    x2 = w - dx2;
                    dy1 = h * a1 / cnstVal2;
                    y1 = vc - dy1;
                    y2 = vc + dy1;
                    x1 = dy1 * dx2 / hd2;
                    let d_val = `M${0},${y1} L${x2},${y1} L${x2},${0} L${w},${vc} L${x2},${h} L${x2},${y2} L${0},${y2} L${x1},${vc} z`;
                    result += `<path d='${d_val}' fill='${(!imgFillFlg ? (grndFillFlg ? "url(#linGrd_" + shpId + ")" : fillColor) : "url(#imgPtrn_" + shpId + ")")}' stroke='${border.color}' stroke-width='${border.width}' stroke-dasharray='${border.strokeDasharray}' />`;
                    break;
                }
                case "homePlate": {
                    const shapAdjst = PPTXXmlUtils.getTextByPathList(node, ["p:spPr", "a:prstGeom", "a:avLst", "a:gd", "attrs", "fmla"]);
                    let adj = 50000 * SLIDE_FACTOR$1;
                    const cnstVal1 = 100000 * SLIDE_FACTOR$1;
                    if (shapAdjst !== undefined) {
                        adj = parseInt(shapAdjst.substr(4)) * SLIDE_FACTOR$1;
                    }
                    var a, x1, dx1, maxAdj, vc = h / 2;
                    const minWH = Math.min(w, h);
                    maxAdj = cnstVal1 * w / minWH;
                    if (adj < 0)
                        a = 0;
                    else if (adj > maxAdj)
                        a = maxAdj;
                    else
                        a = adj;
                    dx1 = minWH * a / cnstVal1;
                    x1 = w - dx1;
                    let d_val = `M${0},${0} L${x1},${0} L${w},${vc} L${x1},${h} L${0},${h} z`;
                    result += `<path  d='${d_val}' fill='${(!imgFillFlg ? (grndFillFlg ? "url(#linGrd_" + shpId + ")" : fillColor) : "url(#imgPtrn_" + shpId + ")")}' stroke='${border.color}' stroke-width='${border.width}' stroke-dasharray='${border.strokeDasharray}' />`;
                    break;
                }
                case "chevron": {
                    const shapAdjst = PPTXXmlUtils.getTextByPathList(node, ["p:spPr", "a:prstGeom", "a:avLst", "a:gd", "attrs", "fmla"]);
                    let adj = 50000 * SLIDE_FACTOR$1;
                    const cnstVal1 = 100000 * SLIDE_FACTOR$1;
                    if (shapAdjst !== undefined) {
                        adj = parseInt(shapAdjst.substr(4)) * SLIDE_FACTOR$1;
                    }
                    var a, x1, dx1, x2, maxAdj, vc = h / 2;
                    const minWH = Math.min(w, h);
                    maxAdj = cnstVal1 * w / minWH;
                    if (adj < 0)
                        a = 0;
                    else if (adj > maxAdj)
                        a = maxAdj;
                    else
                        a = adj;
                    x1 = minWH * a / cnstVal1;
                    x2 = w - x1;
                    let d_val = `M${0},${0} L${x2},${0} L${w},${vc} L${x2},${h} L${0},${h} L${x1},${vc} z`;
                    result += `<path d='${d_val}' fill='${(!imgFillFlg ? (grndFillFlg ? "url(#linGrd_" + shpId + ")" : fillColor) : "url(#imgPtrn_" + shpId + ")")}' stroke='${border.color}' stroke-width='${border.width}' stroke-dasharray='${border.strokeDasharray}' />`;
                    break;
                }
                case "rightArrowCallout": {
                    const shapAdjst_ary = PPTXXmlUtils.getTextByPathList(node, ["p:spPr", "a:prstGeom", "a:avLst", "a:gd"]);
                    let sAdj1, adj1 = 25000 * SLIDE_FACTOR$1;
                    let sAdj2, adj2 = 25000 * SLIDE_FACTOR$1;
                    let sAdj3, adj3 = 25000 * SLIDE_FACTOR$1;
                    let sAdj4, adj4 = 64977 * SLIDE_FACTOR$1;
                    const cnstVal1 = 50000 * SLIDE_FACTOR$1;
                    const cnstVal2 = 100000 * SLIDE_FACTOR$1;
                    const cnstVal3 = 200000 * SLIDE_FACTOR$1;
                    if (shapAdjst_ary !== undefined) {
                        for (const i of shapAdjst_ary.keys()) {
                            const sAdj_name = PPTXXmlUtils.getTextByPathList(shapAdjst_ary[i], ["attrs", "name"]);
                            if (sAdj_name == "adj1") {
                                sAdj1 = PPTXXmlUtils.getTextByPathList(shapAdjst_ary[i], ["attrs", "fmla"]);
                                adj1 = parseInt(sAdj1.substr(4)) * SLIDE_FACTOR$1;
                            }
                            else if (sAdj_name == "adj2") {
                                sAdj2 = PPTXXmlUtils.getTextByPathList(shapAdjst_ary[i], ["attrs", "fmla"]);
                                adj2 = parseInt(sAdj2.substr(4)) * SLIDE_FACTOR$1;
                            }
                            else if (sAdj_name == "adj3") {
                                sAdj3 = PPTXXmlUtils.getTextByPathList(shapAdjst_ary[i], ["attrs", "fmla"]);
                                adj3 = parseInt(sAdj3.substr(4)) * SLIDE_FACTOR$1;
                            }
                            else if (sAdj_name == "adj4") {
                                sAdj4 = PPTXXmlUtils.getTextByPathList(shapAdjst_ary[i], ["attrs", "fmla"]);
                                adj4 = parseInt(sAdj4.substr(4)) * SLIDE_FACTOR$1;
                            }
                        }
                    }
                    var maxAdj2, a2, maxAdj1, a1, maxAdj3, a3, q2, maxAdj4, a4, dy1, dy2, y1, y2, y3, y4, dx3, x3, x2, x1;
                    let vc = h / 2, r = w, b = h, l = 0, t = 0;
                    const ss = Math.min(w, h);
                    maxAdj2 = cnstVal1 * h / ss;
                    a2 = (adj2 < 0) ? 0 : (adj2 > maxAdj2) ? maxAdj2 : adj2;
                    maxAdj1 = a2 * 2;
                    a1 = (adj1 < 0) ? 0 : (adj1 > maxAdj1) ? maxAdj1 : adj1;
                    maxAdj3 = cnstVal2 * w / ss;
                    a3 = (adj3 < 0) ? 0 : (adj3 > maxAdj3) ? maxAdj3 : adj3;
                    q2 = a3 * ss / w;
                    maxAdj4 = cnstVal - q2;
                    a4 = (adj4 < 0) ? 0 : (adj4 > maxAdj4) ? maxAdj4 : adj4;
                    dy1 = ss * a2 / cnstVal2;
                    dy2 = ss * a1 / cnstVal3;
                    y1 = vc - dy1;
                    y2 = vc - dy2;
                    y3 = vc + dy2;
                    y4 = vc + dy1;
                    dx3 = ss * a3 / cnstVal2;
                    x3 = r - dx3;
                    x2 = w * a4 / cnstVal2;
                    x1 = x2 / 2;
                    let d_val = `M${l},${t} L${x2},${t} L${x2},${y2} L${x3},${y2} L${x3},${y1} L${r},${vc} L${x3},${y4} L${x3},${y3} L${x2},${y3} L${x2},${b} L${l},${b} z`;
                    result += `<path d='${d_val}' fill='${(!imgFillFlg ? (grndFillFlg ? "url(#linGrd_" + shpId + ")" : fillColor) : "url(#imgPtrn_" + shpId + ")")}' stroke='${border.color}' stroke-width='${border.width}' stroke-dasharray='${border.strokeDasharray}' />`;
                    break;
                }
                case "downArrowCallout": {
                    const shapAdjst_ary = PPTXXmlUtils.getTextByPathList(node, ["p:spPr", "a:prstGeom", "a:avLst", "a:gd"]);
                    let sAdj1, adj1 = 25000 * SLIDE_FACTOR$1;
                    let sAdj2, adj2 = 25000 * SLIDE_FACTOR$1;
                    let sAdj3, adj3 = 25000 * SLIDE_FACTOR$1;
                    let sAdj4, adj4 = 64977 * SLIDE_FACTOR$1;
                    const cnstVal1 = 50000 * SLIDE_FACTOR$1;
                    const cnstVal2 = 100000 * SLIDE_FACTOR$1;
                    const cnstVal3 = 200000 * SLIDE_FACTOR$1;
                    if (shapAdjst_ary !== undefined) {
                        for (const i of shapAdjst_ary.keys()) {
                            const sAdj_name = PPTXXmlUtils.getTextByPathList(shapAdjst_ary[i], ["attrs", "name"]);
                            if (sAdj_name == "adj1") {
                                sAdj1 = PPTXXmlUtils.getTextByPathList(shapAdjst_ary[i], ["attrs", "fmla"]);
                                adj1 = parseInt(sAdj1.substr(4)) * SLIDE_FACTOR$1;
                            }
                            else if (sAdj_name == "adj2") {
                                sAdj2 = PPTXXmlUtils.getTextByPathList(shapAdjst_ary[i], ["attrs", "fmla"]);
                                adj2 = parseInt(sAdj2.substr(4)) * SLIDE_FACTOR$1;
                            }
                            else if (sAdj_name == "adj3") {
                                sAdj3 = PPTXXmlUtils.getTextByPathList(shapAdjst_ary[i], ["attrs", "fmla"]);
                                adj3 = parseInt(sAdj3.substr(4)) * SLIDE_FACTOR$1;
                            }
                            else if (sAdj_name == "adj4") {
                                sAdj4 = PPTXXmlUtils.getTextByPathList(shapAdjst_ary[i], ["attrs", "fmla"]);
                                adj4 = parseInt(sAdj4.substr(4)) * SLIDE_FACTOR$1;
                            }
                        }
                    }
                    var maxAdj2, a2, maxAdj1, a1, maxAdj3, a3, q2, maxAdj4, a4, dx1, dx2, x1, x2, x3, x4, dy3, y3, y2, y1;
                    let hc = w / 2, r = w, b = h, l = 0, t = 0;
                    const ss = Math.min(w, h);
                    maxAdj2 = cnstVal1 * w / ss;
                    a2 = (adj2 < 0) ? 0 : (adj2 > maxAdj2) ? maxAdj2 : adj2;
                    maxAdj1 = a2 * 2;
                    a1 = (adj1 < 0) ? 0 : (adj1 > maxAdj1) ? maxAdj1 : adj1;
                    maxAdj3 = cnstVal2 * h / ss;
                    a3 = (adj3 < 0) ? 0 : (adj3 > maxAdj3) ? maxAdj3 : adj3;
                    q2 = a3 * ss / h;
                    maxAdj4 = cnstVal2 - q2;
                    a4 = (adj4 < 0) ? 0 : (adj4 > maxAdj4) ? maxAdj4 : adj4;
                    dx1 = ss * a2 / cnstVal2;
                    dx2 = ss * a1 / cnstVal3;
                    x1 = hc - dx1;
                    x2 = hc - dx2;
                    x3 = hc + dx2;
                    x4 = hc + dx1;
                    dy3 = ss * a3 / cnstVal2;
                    y3 = b - dy3;
                    y2 = h * a4 / cnstVal2;
                    y1 = y2 / 2;
                    let d_val = `M${l},${t} L${r},${t} L${r},${y2} L${x3},${y2} L${x3},${y3} L${x4},${y3} L${hc},${b} L${x1},${y3} L${x2},${y3} L${x2},${y2} L${l},${y2} z`;
                    result += `<path d='${d_val}' fill='${(!imgFillFlg ? (grndFillFlg ? "url(#linGrd_" + shpId + ")" : fillColor) : "url(#imgPtrn_" + shpId + ")")}' stroke='${border.color}' stroke-width='${border.width}' stroke-dasharray='${border.strokeDasharray}' />`;
                    break;
                }
                case "leftArrowCallout": {
                    const shapAdjst_ary = PPTXXmlUtils.getTextByPathList(node, ["p:spPr", "a:prstGeom", "a:avLst", "a:gd"]);
                    let sAdj1, adj1 = 25000 * SLIDE_FACTOR$1;
                    let sAdj2, adj2 = 25000 * SLIDE_FACTOR$1;
                    let sAdj3, adj3 = 25000 * SLIDE_FACTOR$1;
                    let sAdj4, adj4 = 64977 * SLIDE_FACTOR$1;
                    const cnstVal1 = 50000 * SLIDE_FACTOR$1;
                    const cnstVal2 = 100000 * SLIDE_FACTOR$1;
                    const cnstVal3 = 200000 * SLIDE_FACTOR$1;
                    if (shapAdjst_ary !== undefined) {
                        for (const i of shapAdjst_ary.keys()) {
                            const sAdj_name = PPTXXmlUtils.getTextByPathList(shapAdjst_ary[i], ["attrs", "name"]);
                            if (sAdj_name == "adj1") {
                                sAdj1 = PPTXXmlUtils.getTextByPathList(shapAdjst_ary[i], ["attrs", "fmla"]);
                                adj1 = parseInt(sAdj1.substr(4)) * SLIDE_FACTOR$1;
                            }
                            else if (sAdj_name == "adj2") {
                                sAdj2 = PPTXXmlUtils.getTextByPathList(shapAdjst_ary[i], ["attrs", "fmla"]);
                                adj2 = parseInt(sAdj2.substr(4)) * SLIDE_FACTOR$1;
                            }
                            else if (sAdj_name == "adj3") {
                                sAdj3 = PPTXXmlUtils.getTextByPathList(shapAdjst_ary[i], ["attrs", "fmla"]);
                                adj3 = parseInt(sAdj3.substr(4)) * SLIDE_FACTOR$1;
                            }
                            else if (sAdj_name == "adj4") {
                                sAdj4 = PPTXXmlUtils.getTextByPathList(shapAdjst_ary[i], ["attrs", "fmla"]);
                                adj4 = parseInt(sAdj4.substr(4)) * SLIDE_FACTOR$1;
                            }
                        }
                    }
                    var maxAdj2, a2, maxAdj1, a1, maxAdj3, a3, q2, maxAdj4, a4, dy1, dy2, y1, y2, y3, y4, x1, dx2, x2, x3;
                    let vc = h / 2, r = w, b = h, l = 0, t = 0;
                    const ss = Math.min(w, h);
                    maxAdj2 = cnstVal1 * h / ss;
                    a2 = (adj2 < 0) ? 0 : (adj2 > maxAdj2) ? maxAdj2 : adj2;
                    maxAdj1 = a2 * 2;
                    a1 = (adj1 < 0) ? 0 : (adj1 > maxAdj1) ? maxAdj1 : adj1;
                    maxAdj3 = cnstVal2 * w / ss;
                    a3 = (adj3 < 0) ? 0 : (adj3 > maxAdj3) ? maxAdj3 : adj3;
                    q2 = a3 * ss / w;
                    maxAdj4 = cnstVal2 - q2;
                    a4 = (adj4 < 0) ? 0 : (adj4 > maxAdj4) ? maxAdj4 : adj4;
                    dy1 = ss * a2 / cnstVal2;
                    dy2 = ss * a1 / cnstVal3;
                    y1 = vc - dy1;
                    y2 = vc - dy2;
                    y3 = vc + dy2;
                    y4 = vc + dy1;
                    x1 = ss * a3 / cnstVal2;
                    dx2 = w * a4 / cnstVal2;
                    x2 = r - dx2;
                    x3 = (x2 + r) / 2;
                    let d_val = `M${l},${vc} L${x1},${y1} L${x1},${y2} L${x2},${y2} L${x2},${t} L${r},${t} L${r},${b} L${x2},${b} L${x2},${y3} L${x1},${y3} L${x1},${y4} z`;
                    result += `<path d='${d_val}' fill='${(!imgFillFlg ? (grndFillFlg ? "url(#linGrd_" + shpId + ")" : fillColor) : "url(#imgPtrn_" + shpId + ")")}' stroke='${border.color}' stroke-width='${border.width}' stroke-dasharray='${border.strokeDasharray}' />`;
                    break;
                }
                case "upArrowCallout": {
                    const shapAdjst_ary = PPTXXmlUtils.getTextByPathList(node, ["p:spPr", "a:prstGeom", "a:avLst", "a:gd"]);
                    let sAdj1, adj1 = 25000 * SLIDE_FACTOR$1;
                    let sAdj2, adj2 = 25000 * SLIDE_FACTOR$1;
                    let sAdj3, adj3 = 25000 * SLIDE_FACTOR$1;
                    let sAdj4, adj4 = 64977 * SLIDE_FACTOR$1;
                    const cnstVal1 = 50000 * SLIDE_FACTOR$1;
                    const cnstVal2 = 100000 * SLIDE_FACTOR$1;
                    const cnstVal3 = 200000 * SLIDE_FACTOR$1;
                    if (shapAdjst_ary !== undefined) {
                        for (const i of shapAdjst_ary.keys()) {
                            const sAdj_name = PPTXXmlUtils.getTextByPathList(shapAdjst_ary[i], ["attrs", "name"]);
                            if (sAdj_name == "adj1") {
                                sAdj1 = PPTXXmlUtils.getTextByPathList(shapAdjst_ary[i], ["attrs", "fmla"]);
                                adj1 = parseInt(sAdj1.substr(4)) * SLIDE_FACTOR$1;
                            }
                            else if (sAdj_name == "adj2") {
                                sAdj2 = PPTXXmlUtils.getTextByPathList(shapAdjst_ary[i], ["attrs", "fmla"]);
                                adj2 = parseInt(sAdj2.substr(4)) * SLIDE_FACTOR$1;
                            }
                            else if (sAdj_name == "adj3") {
                                sAdj3 = PPTXXmlUtils.getTextByPathList(shapAdjst_ary[i], ["attrs", "fmla"]);
                                adj3 = parseInt(sAdj3.substr(4)) * SLIDE_FACTOR$1;
                            }
                            else if (sAdj_name == "adj4") {
                                sAdj4 = PPTXXmlUtils.getTextByPathList(shapAdjst_ary[i], ["attrs", "fmla"]);
                                adj4 = parseInt(sAdj4.substr(4)) * SLIDE_FACTOR$1;
                            }
                        }
                    }
                    var maxAdj2, a2, maxAdj1, a1, maxAdj3, a3, q2, maxAdj4, a4, dx1, dx2, x1, x2, x3, x4, y1, dy2, y2, y3;
                    let hc = w / 2, r = w, b = h, l = 0, t = 0;
                    const ss = Math.min(w, h);
                    maxAdj2 = cnstVal1 * w / ss;
                    a2 = (adj2 < 0) ? 0 : (adj2 > maxAdj2) ? maxAdj2 : adj2;
                    maxAdj1 = a2 * 2;
                    a1 = (adj1 < 0) ? 0 : (adj1 > maxAdj1) ? maxAdj1 : adj1;
                    maxAdj3 = cnstVal2 * h / ss;
                    a3 = (adj3 < 0) ? 0 : (adj3 > maxAdj3) ? maxAdj3 : adj3;
                    q2 = a3 * ss / h;
                    maxAdj4 = cnstVal2 - q2;
                    a4 = (adj4 < 0) ? 0 : (adj4 > maxAdj4) ? maxAdj4 : adj4;
                    dx1 = ss * a2 / cnstVal2;
                    dx2 = ss * a1 / cnstVal3;
                    x1 = hc - dx1;
                    x2 = hc - dx2;
                    x3 = hc + dx2;
                    x4 = hc + dx1;
                    y1 = ss * a3 / cnstVal2;
                    dy2 = h * a4 / cnstVal2;
                    y2 = b - dy2;
                    y3 = (y2 + b) / 2;
                    let d_val = `M${l},${y2} L${x2},${y2} L${x2},${y1} L${x1},${y1} L${hc},${t} L${x4},${y1} L${x3},${y1} L${x3},${y2} L${r},${y2} L${r},${b} L${l},${b} z`;
                    result += `<path d='${d_val}' fill='${(!imgFillFlg ? (grndFillFlg ? "url(#linGrd_" + shpId + ")" : fillColor) : "url(#imgPtrn_" + shpId + ")")}' stroke='${border.color}' stroke-width='${border.width}' stroke-dasharray='${border.strokeDasharray}' />`;
                    break;
                }
                case "leftRightArrowCallout": {
                    const shapAdjst_ary = PPTXXmlUtils.getTextByPathList(node, ["p:spPr", "a:prstGeom", "a:avLst", "a:gd"]);
                    let sAdj1, adj1 = 25000 * SLIDE_FACTOR$1;
                    let sAdj2, adj2 = 25000 * SLIDE_FACTOR$1;
                    let sAdj3, adj3 = 25000 * SLIDE_FACTOR$1;
                    let sAdj4, adj4 = 48123 * SLIDE_FACTOR$1;
                    const cnstVal1 = 50000 * SLIDE_FACTOR$1;
                    const cnstVal2 = 100000 * SLIDE_FACTOR$1;
                    const cnstVal3 = 200000 * SLIDE_FACTOR$1;
                    if (shapAdjst_ary !== undefined) {
                        for (const i of shapAdjst_ary.keys()) {
                            const sAdj_name = PPTXXmlUtils.getTextByPathList(shapAdjst_ary[i], ["attrs", "name"]);
                            if (sAdj_name == "adj1") {
                                sAdj1 = PPTXXmlUtils.getTextByPathList(shapAdjst_ary[i], ["attrs", "fmla"]);
                                adj1 = parseInt(sAdj1.substr(4)) * SLIDE_FACTOR$1;
                            }
                            else if (sAdj_name == "adj2") {
                                sAdj2 = PPTXXmlUtils.getTextByPathList(shapAdjst_ary[i], ["attrs", "fmla"]);
                                adj2 = parseInt(sAdj2.substr(4)) * SLIDE_FACTOR$1;
                            }
                            else if (sAdj_name == "adj3") {
                                sAdj3 = PPTXXmlUtils.getTextByPathList(shapAdjst_ary[i], ["attrs", "fmla"]);
                                adj3 = parseInt(sAdj3.substr(4)) * SLIDE_FACTOR$1;
                            }
                            else if (sAdj_name == "adj4") {
                                sAdj4 = PPTXXmlUtils.getTextByPathList(shapAdjst_ary[i], ["attrs", "fmla"]);
                                adj4 = parseInt(sAdj4.substr(4)) * SLIDE_FACTOR$1;
                            }
                        }
                    }
                    var maxAdj2, a2, maxAdj1, a1, maxAdj3, a3, q2, maxAdj4, a4, dy1, dy2, y1, y2, y3, y4, x1, x4, dx2, x2, x3;
                    let vc = h / 2, hc = w / 2, r = w, b = h, l = 0, t = 0;
                    const ss = Math.min(w, h);
                    maxAdj2 = cnstVal1 * h / ss;
                    a2 = (adj2 < 0) ? 0 : (adj2 > maxAdj2) ? maxAdj2 : adj2;
                    maxAdj1 = a2 * 2;
                    a1 = (adj1 < 0) ? 0 : (adj1 > maxAdj1) ? maxAdj1 : adj1;
                    maxAdj3 = cnstVal1 * w / ss;
                    a3 = (adj3 < 0) ? 0 : (adj3 > maxAdj3) ? maxAdj3 : adj3;
                    q2 = a3 * ss / wd2;
                    maxAdj4 = cnstVal2 - q2;
                    a4 = (adj4 < 0) ? 0 : (adj4 > maxAdj4) ? maxAdj4 : adj4;
                    dy1 = ss * a2 / cnstVal2;
                    dy2 = ss * a1 / cnstVal3;
                    y1 = vc - dy1;
                    y2 = vc - dy2;
                    y3 = vc + dy2;
                    y4 = vc + dy1;
                    x1 = ss * a3 / cnstVal2;
                    x4 = r - x1;
                    dx2 = w * a4 / cnstVal3;
                    x2 = hc - dx2;
                    x3 = hc + dx2;
                    let d_val = `M${l},${vc} L${x1},${y1} L${x1},${y2} L${x2},${y2} L${x2},${t} L${x3},${t} L${x3},${y2} L${x4},${y2} L${x4},${y1} L${r},${vc} L${x4},${y4} L${x4},${y3} L${x3},${y3} L${x3},${b} L${x2},${b} L${x2},${y3} L${x1},${y3} L${x1},${y4} z`;
                    result += `<path d='${d_val}' fill='${(!imgFillFlg ? (grndFillFlg ? "url(#linGrd_" + shpId + ")" : fillColor) : "url(#imgPtrn_" + shpId + ")")}' stroke='${border.color}' stroke-width='${border.width}' stroke-dasharray='${border.strokeDasharray}' />`;
                    break;
                }
                case "quadArrowCallout": {
                    const shapAdjst_ary = PPTXXmlUtils.getTextByPathList(node, ["p:spPr", "a:prstGeom", "a:avLst", "a:gd"]);
                    let sAdj1, adj1 = 18515 * SLIDE_FACTOR$1;
                    let sAdj2, adj2 = 18515 * SLIDE_FACTOR$1;
                    let sAdj3, adj3 = 18515 * SLIDE_FACTOR$1;
                    let sAdj4, adj4 = 48123 * SLIDE_FACTOR$1;
                    const cnstVal1 = 50000 * SLIDE_FACTOR$1;
                    const cnstVal2 = 100000 * SLIDE_FACTOR$1;
                    const cnstVal3 = 200000 * SLIDE_FACTOR$1;
                    if (shapAdjst_ary !== undefined) {
                        for (const i of shapAdjst_ary.keys()) {
                            const sAdj_name = PPTXXmlUtils.getTextByPathList(shapAdjst_ary[i], ["attrs", "name"]);
                            if (sAdj_name == "adj1") {
                                sAdj1 = PPTXXmlUtils.getTextByPathList(shapAdjst_ary[i], ["attrs", "fmla"]);
                                adj1 = parseInt(sAdj1.substr(4)) * SLIDE_FACTOR$1;
                            }
                            else if (sAdj_name == "adj2") {
                                sAdj2 = PPTXXmlUtils.getTextByPathList(shapAdjst_ary[i], ["attrs", "fmla"]);
                                adj2 = parseInt(sAdj2.substr(4)) * SLIDE_FACTOR$1;
                            }
                            else if (sAdj_name == "adj3") {
                                sAdj3 = PPTXXmlUtils.getTextByPathList(shapAdjst_ary[i], ["attrs", "fmla"]);
                                adj3 = parseInt(sAdj3.substr(4)) * SLIDE_FACTOR$1;
                            }
                            else if (sAdj_name == "adj4") {
                                sAdj4 = PPTXXmlUtils.getTextByPathList(shapAdjst_ary[i], ["attrs", "fmla"]);
                                adj4 = parseInt(sAdj4.substr(4)) * SLIDE_FACTOR$1;
                            }
                        }
                    }
                    let vc = h / 2, hc = w / 2, r = w, b = h, l = 0, t = 0;
                    const ss = Math.min(w, h);
                    var a2, maxAdj1, a1, maxAdj3, a3, q2, maxAdj4, a4, dx2, dx3, ah, dx1, dy1, x8, x2, x7, x3, x6, x4, x5, y8, y2, y7, y3, y6, y4, y5;
                    a2 = (adj2 < 0) ? 0 : (adj2 > cnstVal1) ? cnstVal1 : adj2;
                    maxAdj1 = a2 * 2;
                    a1 = (adj1 < 0) ? 0 : (adj1 > maxAdj1) ? maxAdj1 : adj1;
                    maxAdj3 = cnstVal1 - a2;
                    a3 = (adj3 < 0) ? 0 : (adj3 > maxAdj3) ? maxAdj3 : adj3;
                    q2 = a3 * 2;
                    maxAdj4 = cnstVal2 - q2;
                    a4 = (adj4 < a1) ? a1 : (adj4 > maxAdj4) ? maxAdj4 : adj4;
                    dx2 = ss * a2 / cnstVal2;
                    dx3 = ss * a1 / cnstVal3;
                    ah = ss * a3 / cnstVal2;
                    dx1 = w * a4 / cnstVal3;
                    dy1 = h * a4 / cnstVal3;
                    x8 = r - ah;
                    x2 = hc - dx1;
                    x7 = hc + dx1;
                    x3 = hc - dx2;
                    x6 = hc + dx2;
                    x4 = hc - dx3;
                    x5 = hc + dx3;
                    y8 = b - ah;
                    y2 = vc - dy1;
                    y7 = vc + dy1;
                    y3 = vc - dx2;
                    y6 = vc + dx2;
                    y4 = vc - dx3;
                    y5 = vc + dx3;
                    let d_val = `M${l},${vc} L${ah},${y3} L${ah},${y4} L${x2},${y4} L${x2},${y2} L${x4},${y2} L${x4},${ah} L${x3},${ah} L${hc},${t} L${x6},${ah} L${x5},${ah} L${x5},${y2} L${x7},${y2} L${x7},${y4} L${x8},${y4} L${x8},${y3} L${r},${vc} L${x8},${y6} L${x8},${y5} L${x7},${y5} L${x7},${y7} L${x5},${y7} L${x5},${y8} L${x6},${y8} L${hc},${b} L${x3},${y8} L${x4},${y8} L${x4},${y7} L${x2},${y7} L${x2},${y5} L${ah},${y5} L${ah},${y6} z`;
                    result += `<path d='${d_val}' fill='${(!imgFillFlg ? (grndFillFlg ? "url(#linGrd_" + shpId + ")" : fillColor) : "url(#imgPtrn_" + shpId + ")")}' stroke='${border.color}' stroke-width='${border.width}' stroke-dasharray='${border.strokeDasharray}' />`;
                    break;
                }
                case "curvedDownArrow": {
                    const shapAdjst_ary = PPTXXmlUtils.getTextByPathList(node, ["p:spPr", "a:prstGeom", "a:avLst", "a:gd"]);
                    let sAdj1, adj1 = 25000 * SLIDE_FACTOR$1;
                    let sAdj2, adj2 = 50000 * SLIDE_FACTOR$1;
                    let sAdj3, adj3 = 25000 * SLIDE_FACTOR$1;
                    const cnstVal1 = 50000 * SLIDE_FACTOR$1;
                    const cnstVal2 = 100000 * SLIDE_FACTOR$1;
                    if (shapAdjst_ary !== undefined) {
                        for (const i of shapAdjst_ary.keys()) {
                            const sAdj_name = PPTXXmlUtils.getTextByPathList(shapAdjst_ary[i], ["attrs", "name"]);
                            if (sAdj_name == "adj1") {
                                sAdj1 = PPTXXmlUtils.getTextByPathList(shapAdjst_ary[i], ["attrs", "fmla"]);
                                adj1 = parseInt(sAdj1.substr(4)) * SLIDE_FACTOR$1;
                            }
                            else if (sAdj_name == "adj2") {
                                sAdj2 = PPTXXmlUtils.getTextByPathList(shapAdjst_ary[i], ["attrs", "fmla"]);
                                adj2 = parseInt(sAdj2.substr(4)) * SLIDE_FACTOR$1;
                            }
                            else if (sAdj_name == "adj3") {
                                sAdj3 = PPTXXmlUtils.getTextByPathList(shapAdjst_ary[i], ["attrs", "fmla"]);
                                adj3 = parseInt(sAdj3.substr(4)) * SLIDE_FACTOR$1;
                            }
                        }
                    }
                    let cw = (drawW !== undefined) ? drawW : w;
                    let ch = (drawH !== undefined) ? drawH : h;
                    var vc = ch / 2, hc = cw / 2, wd2 = cw / 2, r = cw, b = ch, l = 0, t = 0, c3d4 = 270, cd2 = 180, cd4 = 90;
                    const ss = Math.min(cw, ch);
                    var maxAdj2, a2, a1, th, aw, q1, wR, q7, q8, q9, q10, q11, idy, maxAdj3, a3, ah, x3, q2, q3, q4, q5, dx, x5, x7, q6, dh, x4, x8, aw2, x6, y1, swAng, mswAng, q12, dang2, stAng, stAng2, swAng2, swAng3;
                    function fmt(num) {
                        return parseFloat(num.toFixed(2));
                    }
                    maxAdj2 = cnstVal1 * cw / ss;
                    a2 = (adj2 < 0) ? 0 : (adj2 > maxAdj2) ? maxAdj2 : adj2;
                    a1 = (adj1 < 0) ? 0 : (adj1 > cnstVal2) ? cnstVal2 : adj1;
                    th = ss * a1 / cnstVal2;
                    aw = ss * a2 / cnstVal2;
                    q1 = (th + aw) / 4;
                    wR = wd2 - q1;
                    q7 = wR * 2;
                    q8 = q7 * q7;
                    q9 = th * th;
                    q10 = q8 - q9;
                    q11 = Math.sqrt(q10);
                    idy = q11 * ch / q7;
                    maxAdj3 = cnstVal2 * idy / ss;
                    a3 = (adj3 < 0) ? 0 : (adj3 > maxAdj3) ? maxAdj3 : adj3;
                    ah = ss * adj3 / cnstVal2;
                    x3 = wR + th;
                    q2 = ch * ch;
                    q3 = ah * ah;
                    q4 = q2 - q3;
                    q5 = Math.sqrt(q4);
                    dx = q5 * wR / ch;
                    x5 = wR + dx;
                    x7 = x3 + dx;
                    q6 = aw - th;
                    dh = q6 / 2;
                    x4 = x5 - dh;
                    x8 = x7 + dh;
                    aw2 = aw / 2;
                    x6 = r - aw2;
                    y1 = b - ah;
                    swAng = Math.atan(dx / ah);
                    const swAngDeg = swAng * 180 / Math.PI;
                    mswAng = -swAngDeg;
                    q12 = th / 2;
                    dang2 = Math.atan(q12 / idy);
                    const dang2Deg = dang2 * 180 / Math.PI;
                    stAng = c3d4 + swAngDeg;
                    stAng2 = c3d4 - dang2Deg;
                    swAng2 = dang2Deg - cd4;
                    swAng3 = cd4 + dang2Deg;
                    x6 = fmt(x6);
                    b = fmt(b);
                    x4 = fmt(x4);
                    y1 = fmt(y1);
                    x5 = fmt(x5);
                    x3 = fmt(x3);
                    t = fmt(t);
                    th = fmt(th);
                    x8 = fmt(x8);
                    wR = fmt(wR);
                    ch = fmt(ch);
                    let d_val = `M${x6},${b} L${x4},${y1} L${x5},${y1}${PPTXShapeUtils.shapeArc(wR, ch, wR, ch, stAng, (stAng + mswAng), false).replace("M", "L")} L${x3},${t}${PPTXShapeUtils.shapeArc(x3, ch, wR, ch, c3d4, (c3d4 + swAngDeg), false).replace("M", "L")} L${fmt(x5 + th)},${y1} L${x8},${y1} zM${x3},${t}${PPTXShapeUtils.shapeArc(x3, ch, wR, ch, stAng2, (stAng2 + swAng2), false).replace("M", "L")}${PPTXShapeUtils.shapeArc(wR, ch, wR, ch, cd2, (cd2 + swAng3), false).replace("M", "L")}`;
                    result += `<path d='${d_val}' fill='${(!imgFillFlg ? (grndFillFlg ? "url(#linGrd_" + shpId + ")" : fillColor) : "url(#imgPtrn_" + shpId + ")")}' stroke='${border.color}' stroke-width='${border.width}' stroke-dasharray='${border.strokeDasharray}' />`;
                    break;
                }
                case "curvedLeftArrow": {
                    const shapAdjst_ary = PPTXXmlUtils.getTextByPathList(node, ["p:spPr", "a:prstGeom", "a:avLst", "a:gd"]);
                    let sAdj1, adj1 = 25000 * SLIDE_FACTOR$1;
                    let sAdj2, adj2 = 50000 * SLIDE_FACTOR$1;
                    let sAdj3, adj3 = 25000 * SLIDE_FACTOR$1;
                    const cnstVal1 = 50000 * SLIDE_FACTOR$1;
                    const cnstVal2 = 100000 * SLIDE_FACTOR$1;
                    if (shapAdjst_ary !== undefined) {
                        for (const i of shapAdjst_ary.keys()) {
                            const sAdj_name = PPTXXmlUtils.getTextByPathList(shapAdjst_ary[i], ["attrs", "name"]);
                            if (sAdj_name == "adj1") {
                                sAdj1 = PPTXXmlUtils.getTextByPathList(shapAdjst_ary[i], ["attrs", "fmla"]);
                                adj1 = parseInt(sAdj1.substr(4)) * SLIDE_FACTOR$1;
                            }
                            else if (sAdj_name == "adj2") {
                                sAdj2 = PPTXXmlUtils.getTextByPathList(shapAdjst_ary[i], ["attrs", "fmla"]);
                                adj2 = parseInt(sAdj2.substr(4)) * SLIDE_FACTOR$1;
                            }
                            else if (sAdj_name == "adj3") {
                                sAdj3 = PPTXXmlUtils.getTextByPathList(shapAdjst_ary[i], ["attrs", "fmla"]);
                                adj3 = parseInt(sAdj3.substr(4)) * SLIDE_FACTOR$1;
                            }
                        }
                    }
                    let cw = (drawW !== undefined) ? drawW : w;
                    let ch = (drawH !== undefined) ? drawH : h;
                    var vc = ch / 2, hc = cw / 2, hd2 = ch / 2, r = cw, b = ch, l = 0, t = 0, c3d4 = 270, cd2 = 180, cd4 = 90;
                    const ss = Math.min(cw, ch);
                    var maxAdj2, a2, a1, th, aw, q1, hR, q7, q8, q9, q10, q11, iDx, maxAdj3, a3, ah, y3, q2, q3, q4, q5, dy, y5, y7, q6, dh, y4, y8, aw2, y6, x1, swAng, mswAng, q12, dang2, swAng2, swAng3, stAng3;
                    function fmt(num) {
                        return parseFloat(num.toFixed(2));
                    }
                    maxAdj2 = cnstVal1 * ch / ss;
                    a2 = (adj2 < 0) ? 0 : (adj2 > maxAdj2) ? maxAdj2 : adj2;
                    a1 = (adj1 < 0) ? 0 : (adj1 > a2) ? a2 : adj1;
                    th = ss * a1 / cnstVal2;
                    aw = ss * a2 / cnstVal2;
                    q1 = (th + aw) / 4;
                    hR = hd2 - q1;
                    q7 = hR * 2;
                    q8 = q7 * q7;
                    q9 = th * th;
                    q10 = q8 - q9;
                    q11 = Math.sqrt(q10);
                    iDx = q11 * cw / q7;
                    maxAdj3 = cnstVal2 * iDx / ss;
                    a3 = (adj3 < 0) ? 0 : (adj3 > maxAdj3) ? maxAdj3 : adj3;
                    ah = ss * a3 / cnstVal2;
                    y3 = hR + th;
                    q2 = cw * cw;
                    q3 = ah * ah;
                    q4 = q2 - q3;
                    q5 = Math.sqrt(q4);
                    dy = q5 * hR / cw;
                    y5 = hR + dy;
                    y7 = y3 + dy;
                    q6 = aw - th;
                    dh = q6 / 2;
                    y4 = y5 - dh;
                    y8 = y7 + dh;
                    aw2 = aw / 2;
                    y6 = b - aw2;
                    x1 = l + ah;
                    swAng = Math.atan(dy / ah);
                    mswAng = -swAng;
                    q12 = th / 2;
                    dang2 = Math.atan(q12 / iDx);
                    swAng2 = dang2 - swAng;
                    swAng3 = swAng + dang2;
                    stAng3 = -dang2;
                    var swAngDg, swAng2Dg, stAng3dg;
                    swAngDg = swAng * 180 / Math.PI;
                    swAng2Dg = swAng2 * 180 / Math.PI;
                    stAng3dg = stAng3 * 180 / Math.PI;
                    r = fmt(r);
                    y3 = fmt(y3);
                    l = fmt(l);
                    hR = fmt(hR);
                    cw = fmt(cw);
                    t = fmt(t);
                    x1 = fmt(x1);
                    y7 = fmt(y7);
                    y8 = fmt(y8);
                    y6 = fmt(y6);
                    y4 = fmt(y4);
                    y5 = fmt(y5);
                    let d_val = `M${r},${y3}${PPTXShapeUtils.shapeArc(l, hR, cw, hR, 0, -cd4, false).replace("M", "L")} L${l},${t}${PPTXShapeUtils.shapeArc(l, y3, cw, hR, c3d4, (c3d4 + cd4), false).replace("M", "L")} L${r},${y3}${PPTXShapeUtils.shapeArc(l, y3, cw, hR, 0, swAngDg, false).replace("M", "L")} L${x1},${y7} L${x1},${y8} L${l},${y6} L${x1},${y4} L${x1},${y5}${PPTXShapeUtils.shapeArc(l, hR, cw, hR, swAngDg, (swAngDg + swAng2Dg), false).replace("M", "L")}${PPTXShapeUtils.shapeArc(l, hR, cw, hR, 0, -cd4, false).replace("M", "L")}${PPTXShapeUtils.shapeArc(l, y3, cw, hR, c3d4, (c3d4 + cd4), false).replace("M", "L")}`;
                    result += `<path d='${d_val}' fill='${(!imgFillFlg ? (grndFillFlg ? "url(#linGrd_" + shpId + ")" : fillColor) : "url(#imgPtrn_" + shpId + ")")}' stroke='${border.color}' stroke-width='${border.width}' stroke-dasharray='${border.strokeDasharray}' />`;
                    break;
                }
                case "curvedRightArrow": {
                    const shapAdjst_ary = PPTXXmlUtils.getTextByPathList(node, ["p:spPr", "a:prstGeom", "a:avLst", "a:gd"]);
                    let sAdj1, adj1 = 25000 * SLIDE_FACTOR$1;
                    let sAdj2, adj2 = 50000 * SLIDE_FACTOR$1;
                    let sAdj3, adj3 = 25000 * SLIDE_FACTOR$1;
                    const cnstVal1 = 50000 * SLIDE_FACTOR$1;
                    const cnstVal2 = 100000 * SLIDE_FACTOR$1;
                    if (shapAdjst_ary !== undefined) {
                        for (const i of shapAdjst_ary.keys()) {
                            const sAdj_name = PPTXXmlUtils.getTextByPathList(shapAdjst_ary[i], ["attrs", "name"]);
                            if (sAdj_name == "adj1") {
                                sAdj1 = PPTXXmlUtils.getTextByPathList(shapAdjst_ary[i], ["attrs", "fmla"]);
                                adj1 = parseInt(sAdj1.substr(4)) * SLIDE_FACTOR$1;
                            }
                            else if (sAdj_name == "adj2") {
                                sAdj2 = PPTXXmlUtils.getTextByPathList(shapAdjst_ary[i], ["attrs", "fmla"]);
                                adj2 = parseInt(sAdj2.substr(4)) * SLIDE_FACTOR$1;
                            }
                            else if (sAdj_name == "adj3") {
                                sAdj3 = PPTXXmlUtils.getTextByPathList(shapAdjst_ary[i], ["attrs", "fmla"]);
                                adj3 = parseInt(sAdj3.substr(4)) * SLIDE_FACTOR$1;
                            }
                        }
                    }
                    let cw = (drawW !== undefined) ? drawW : w;
                    let ch = (drawH !== undefined) ? drawH : h;
                    var vc = ch / 2, hc = cw / 2, hd2 = ch / 2, r = cw, b = ch, l = 0, t = 0, c3d4 = 270, cd2 = 180, cd4 = 90;
                    const ss = Math.min(cw, ch);
                    var maxAdj2, a2, a1, th, aw, q1, hR, q7, q8, q9, q10, q11, iDx, maxAdj3, a3, ah, y3, q2, q3, q4, q5, dy, y5, y7, q6, dh, y4, y8, aw2, y6, x1, swAng, stAng, mswAng, q12, dang2, swAng2, swAng3, stAng3;
                    maxAdj2 = cnstVal1 * ch / ss;
                    a2 = (adj2 < 0) ? 0 : (adj2 > maxAdj2) ? maxAdj2 : adj2;
                    a1 = (adj1 < 0) ? 0 : (adj1 > a2) ? a2 : adj1;
                    th = ss * a1 / cnstVal2;
                    aw = ss * a2 / cnstVal2;
                    q1 = (th + aw) / 4;
                    hR = hd2 - q1;
                    q7 = hR * 2;
                    q8 = q7 * q7;
                    q9 = th * th;
                    q10 = q8 - q9;
                    q11 = Math.sqrt(q10);
                    iDx = q11 * cw / q7;
                    maxAdj3 = cnstVal2 * iDx / ss;
                    a3 = (adj3 < 0) ? 0 : (adj3 > maxAdj3) ? maxAdj3 : adj3;
                    ah = ss * a3 / cnstVal2;
                    y3 = hR + th;
                    q2 = cw * cw;
                    q3 = ah * ah;
                    q4 = q2 - q3;
                    q5 = Math.sqrt(q4);
                    dy = q5 * hR / cw;
                    y5 = hR + dy;
                    y7 = y3 + dy;
                    q6 = aw - th;
                    dh = q6 / 2;
                    y4 = y5 - dh;
                    y8 = y7 + dh;
                    aw2 = aw / 2;
                    y6 = b - aw2;
                    x1 = r - ah;
                    swAng = Math.atan(dy / ah);
                    stAng = Math.PI + 0 - swAng;
                    mswAng = -swAng;
                    q12 = th / 2;
                    dang2 = Math.atan(q12 / iDx);
                    swAng2 = dang2 - Math.PI / 2;
                    swAng3 = Math.PI / 2 + dang2;
                    stAng3 = Math.PI - dang2;
                    var stAngDg, mswAngDg, swAngDg, swAng2dg;
                    stAngDg = stAng * 180 / Math.PI;
                    mswAngDg = mswAng * 180 / Math.PI;
                    swAngDg = swAng * 180 / Math.PI;
                    swAng2dg = swAng2 * 180 / Math.PI;
                    let d_val = `M${l},${hR}${shapeArcAlt(cw, hR, cw, hR, cd2, cd2 + mswAngDg, false).replace("M", "L")} L${x1},${y5} L${x1},${y4} L${r},${y6} L${x1},${y8} L${x1},${y7}${shapeArcAlt(cw, y3, cw, hR, stAngDg, stAngDg + swAngDg, false).replace("M", "L")} L${l},${hR}${shapeArcAlt(cw, hR, cw, hR, cd2, cd2 + cd4, false).replace("M", "L")} L${r},${th}${shapeArcAlt(cw, y3, cw, hR, c3d4, c3d4 + swAng2dg, false).replace("M", "L")} z`;
                    result += `<path d='${d_val}' fill='${(!imgFillFlg ? (grndFillFlg ? "url(#linGrd_" + shpId + ")" : fillColor) : "url(#imgPtrn_" + shpId + ")")}' stroke='${border.color}' stroke-width='${border.width}' stroke-dasharray='${border.strokeDasharray}' />`;
                    break;
                }
                case "curvedUpArrow": {
                    const shapAdjst_ary = PPTXXmlUtils.getTextByPathList(node, ["p:spPr", "a:prstGeom", "a:avLst", "a:gd"]);
                    let sAdj1, adj1 = 25000 * SLIDE_FACTOR$1;
                    let sAdj2, adj2 = 50000 * SLIDE_FACTOR$1;
                    let sAdj3, adj3 = 25000 * SLIDE_FACTOR$1;
                    const cnstVal1 = 50000 * SLIDE_FACTOR$1;
                    const cnstVal2 = 100000 * SLIDE_FACTOR$1;
                    if (shapAdjst_ary !== undefined) {
                        for (const i of shapAdjst_ary.keys()) {
                            const sAdj_name = PPTXXmlUtils.getTextByPathList(shapAdjst_ary[i], ["attrs", "name"]);
                            if (sAdj_name == "adj1") {
                                sAdj1 = PPTXXmlUtils.getTextByPathList(shapAdjst_ary[i], ["attrs", "fmla"]);
                                adj1 = parseInt(sAdj1.substr(4)) * SLIDE_FACTOR$1;
                            }
                            else if (sAdj_name == "adj2") {
                                sAdj2 = PPTXXmlUtils.getTextByPathList(shapAdjst_ary[i], ["attrs", "fmla"]);
                                adj2 = parseInt(sAdj2.substr(4)) * SLIDE_FACTOR$1;
                            }
                            else if (sAdj_name == "adj3") {
                                sAdj3 = PPTXXmlUtils.getTextByPathList(shapAdjst_ary[i], ["attrs", "fmla"]);
                                adj3 = parseInt(sAdj3.substr(4)) * SLIDE_FACTOR$1;
                            }
                        }
                    }
                    let cw = (drawW !== undefined) ? drawW : w;
                    let ch = (drawH !== undefined) ? drawH : h;
                    var vc = ch / 2, hc = cw / 2, wd2 = cw / 2, r = cw, b = ch, l = 0, t = 0, c3d4 = 270, cd2 = 180, cd4 = 90;
                    const ss = Math.min(cw, ch);
                    var maxAdj2, a2, a1, th, aw, q1, wR, q7, q8, q9, q10, q11, idy, maxAdj3, a3, ah, x3, q2, q3, q4, q5, dx, x5, x7, q6, dh, x4, x8, aw2, x6, y1, swAng, mswAng, q12, dang2, swAng2, stAng3, swAng3, stAng2;
                    function fmt(num) {
                        return parseFloat(num.toFixed(2));
                    }
                    function fmtArc(arcStr) {
                        return arcStr.replace(/[-+]?\d*\.?\d+(?:[eE][-+]?\d+)?/g, (match) => {
                            return fmt(parseFloat(match)).toString();
                        });
                    }
                    maxAdj2 = cnstVal1 * cw / ss;
                    a2 = (adj2 < 0) ? 0 : (adj2 > maxAdj2) ? maxAdj2 : adj2;
                    a1 = (adj1 < 0) ? 0 : (adj1 > cnstVal2) ? cnstVal2 : adj1;
                    th = ss * a1 / cnstVal2;
                    aw = ss * a2 / cnstVal2;
                    q1 = (th + aw) / 4;
                    wR = wd2 - q1;
                    q7 = wR * 2;
                    q8 = q7 * q7;
                    q9 = th * th;
                    q10 = q8 - q9;
                    q11 = Math.sqrt(q10);
                    idy = q11 * ch / q7;
                    maxAdj3 = cnstVal2 * idy / ss;
                    a3 = (adj3 < 0) ? 0 : (adj3 > maxAdj3) ? maxAdj3 : adj3;
                    ah = ss * adj3 / cnstVal2;
                    x3 = wR + th;
                    q2 = ch * ch;
                    q3 = ah * ah;
                    q4 = q2 - q3;
                    q5 = Math.sqrt(q4);
                    dx = q5 * wR / ch;
                    x5 = wR + dx;
                    x7 = x3 + dx;
                    q6 = aw - th;
                    dh = q6 / 2;
                    x4 = x5 - dh;
                    x8 = x7 + dh;
                    aw2 = aw / 2;
                    x6 = r - aw2;
                    y1 = t + ah;
                    swAng = Math.atan(dx / ah);
                    mswAng = -swAng;
                    q12 = th / 2;
                    dang2 = Math.atan(q12 / idy);
                    swAng2 = dang2 - swAng;
                    stAng3 = Math.PI / 2 - swAng;
                    swAng3 = swAng + dang2;
                    stAng2 = Math.PI / 2 - dang2;
                    var stAng2dg, swAng2dg, swAngDg, swAng2dg;
                    stAng2dg = stAng2 * 180 / Math.PI;
                    swAng2dg = swAng2 * 180 / Math.PI;
                    stAng3dg = stAng3 * 180 / Math.PI;
                    swAngDg = swAng * 180 / Math.PI;
                    wR = fmt(wR);
                    ch = fmt(ch);
                    cw = fmt(cw);
                    x3 = fmt(x3);
                    x5 = fmt(x5);
                    x7 = fmt(x7);
                    x4 = fmt(x4);
                    x8 = fmt(x8);
                    x6 = fmt(x6);
                    y1 = fmt(y1);
                    b = fmt(b);
                    th = fmt(th);
                    t = fmt(t);
                    let d_val = `${fmtArc(PPTXShapeUtils.shapeArc(wR, 0, wR, ch, stAng2dg, stAng2dg + swAng2dg, false))} L${x5},${y1} L${x4},${y1} L${x6},${t} L${x8},${y1} L${x7},${y1}${fmtArc(PPTXShapeUtils.shapeArc(x3, 0, wR, ch, stAng3dg, stAng3dg + swAngDg, false)).replace("M", "L")} L${wR},${b}${fmtArc(PPTXShapeUtils.shapeArc(wR, 0, wR, ch, cd4, cd2, false)).replace("M", "L")} L${th},${t}${fmtArc(PPTXShapeUtils.shapeArc(x3, 0, wR, ch, cd2, cd4, false)).replace("M", "L")}`;
                    result += `<path d='${d_val}' fill='${(!imgFillFlg ? (grndFillFlg ? "url(#linGrd_" + shpId + ")" : fillColor) : "url(#imgPtrn_" + shpId + ")")}' stroke='${border.color}' stroke-width='${border.width}' stroke-dasharray='${border.strokeDasharray}' />`;
                    break;
                }
                case "mathDivide":
                case "mathEqual":
                case "mathMinus":
                case "mathMultiply":
                case "mathNotEqual":
                case "mathPlus": {
                    result += renderMathSymbol(shapType, drawW, drawH, imgFillFlg, grndFillFlg, fillColor, border, shpId, node);
                    break;
                }
                case "cylinder":
                case "can":
                case "flowChartMagneticDisk":
                case "flowChartMagneticDrum": {
                    const shapAdjst = PPTXXmlUtils.getTextByPathList(node, ["p:spPr", "a:prstGeom", "a:avLst", "a:gd", "attrs", "fmla"]);
                    let adj = 25000 * SLIDE_FACTOR$1;
                    const cnstVal1 = 50000 * SLIDE_FACTOR$1;
                    const cnstVal2 = 200000 * SLIDE_FACTOR$1;
                    if (shapAdjst !== undefined) {
                        adj = parseInt(shapAdjst.substr(4)) * SLIDE_FACTOR$1;
                    }
                    const ss = Math.min(w, h);
                    var maxAdj, a, y1, y2, y3, dVal;
                    if (shapType == "flowChartMagneticDisk" || shapType == "flowChartMagneticDrum") {
                        adj = 50000 * SLIDE_FACTOR$1;
                    }
                    maxAdj = cnstVal1 * h / ss;
                    a = (adj < 0) ? 0 : (adj > maxAdj) ? maxAdj : adj;
                    y1 = ss * a / cnstVal2;
                    y2 = y1 + y1;
                    y3 = h - y1;
                    var cd2 = 180, wd2 = w / 2;
                    let tranglRott = "";
                    if (shapType == "flowChartMagneticDrum") {
                        tranglRott = `transform='rotate(90 ${w / 2},${h / 2})'`;
                    }
                    dVal = `${shapeArcAlt(wd2, y1, wd2, y1, 0, cd2, false)}${shapeArcAlt(wd2, y1, wd2, y1, cd2, cd2 + cd2, false).replace("M", "L")} L${w},${y3}${shapeArcAlt(wd2, y3, wd2, y1, 0, cd2, false).replace("M", "L")} L${0},${y1}`;
                    result += `<path ${tranglRott} d='${dVal}' fill='${(!imgFillFlg ? (grndFillFlg ? "url(#linGrd_" + shpId + ")" : fillColor) : "url(#imgPtrn_" + shpId + ")")}' stroke='${border.color}' stroke-width='${border.width}' stroke-dasharray='${border.strokeDasharray}' />`;
                    break;
                }
                case "swooshArrow": {
                    const shapAdjst_ary = PPTXXmlUtils.getTextByPathList(node, ["p:spPr", "a:prstGeom", "a:avLst", "a:gd"]);
                    const refr = SLIDE_FACTOR$1;
                    let sAdj1, adj1 = 25000 * refr;
                    let sAdj2, adj2 = 16667 * refr;
                    if (shapAdjst_ary !== undefined) {
                        for (const i of shapAdjst_ary.keys()) {
                            const sAdj_name = PPTXXmlUtils.getTextByPathList(shapAdjst_ary[i], ["attrs", "name"]);
                            if (sAdj_name == "adj1") {
                                sAdj1 = PPTXXmlUtils.getTextByPathList(shapAdjst_ary[i], ["attrs", "fmla"]);
                                adj1 = parseInt(sAdj1.substr(4)) * refr;
                            }
                            else if (sAdj_name == "adj2") {
                                sAdj2 = PPTXXmlUtils.getTextByPathList(shapAdjst_ary[i], ["attrs", "fmla"]);
                                adj2 = parseInt(sAdj2.substr(4)) * refr;
                            }
                        }
                    }
                    const cnstVal1 = 1 * refr;
                    const cnstVal2 = 70000 * refr;
                    const cnstVal3 = 75000 * refr;
                    const cnstVal4 = 100000 * refr;
                    const ss = Math.min(w, h);
                    const ssd8 = ss / 8;
                    const hd6 = h / 6;
                    let a1, maxAdj2, a2, ad1, ad2, xB, yB, alfa, dx0, xC, dx1, yF, xF, xE, yE, dy2, dy22, dy3, yD, dy4, yP1, xP1, dy5, yP2, xP2;
                    a1 = (adj1 < cnstVal1) ? cnstVal1 : (adj1 > cnstVal3) ? cnstVal3 : adj1;
                    maxAdj2 = cnstVal2 * w / ss;
                    a2 = (adj2 < 0) ? 0 : (adj2 > maxAdj2) ? maxAdj2 : adj2;
                    ad1 = h * a1 / cnstVal4;
                    ad2 = ss * a2 / cnstVal4;
                    xB = w - ad2;
                    yB = ssd8;
                    alfa = (Math.PI / 2) / 14;
                    dx0 = ssd8 * Math.tan(alfa);
                    xC = xB - dx0;
                    dx1 = ad1 * Math.tan(alfa);
                    yF = yB + ad1;
                    xF = xB + dx1;
                    xE = xF + dx0;
                    yE = yF + ssd8;
                    dy2 = yE - 0;
                    dy22 = dy2 / 2;
                    dy3 = h / 20;
                    yD = dy22 - dy3;
                    dy4 = hd6;
                    yP1 = hd6 + dy4;
                    xP1 = w / 6;
                    dy5 = hd6 / 2;
                    yP2 = yF + dy5;
                    xP2 = w / 4;
                    let dVal = `M${0},${h} Q${xP1},${yP1} ${xB},${yB} L${xC},${0} L${w},${yD} L${xE},${yE} L${xF},${yF} Q${xP2},${yP2} ${0},${h} z`;
                    result += `<path d='${dVal}' fill='${(!imgFillFlg ? (grndFillFlg ? "url(#linGrd_" + shpId + ")" : fillColor) : "url(#imgPtrn_" + shpId + ")")}' stroke='${border.color}' stroke-width='${border.width}' stroke-dasharray='${border.strokeDasharray}' />`;
                    break;
                }
                case "circularArrow": {
                    const shapAdjst_ary = PPTXXmlUtils.getTextByPathList(node, ["p:spPr", "a:prstGeom", "a:avLst", "a:gd"]);
                    let sAdj1, adj1 = 12500 * SLIDE_FACTOR$1;
                    let sAdj2, adj2 = (1142319 / 60000) * Math.PI / 180;
                    let sAdj3, adj3 = (20457681 / 60000) * Math.PI / 180;
                    let sAdj4, adj4 = (10800000 / 60000) * Math.PI / 180;
                    let sAdj5, adj5 = 12500 * SLIDE_FACTOR$1;
                    if (shapAdjst_ary !== undefined) {
                        for (const i of shapAdjst_ary.keys()) {
                            const sAdj_name = PPTXXmlUtils.getTextByPathList(shapAdjst_ary[i], ["attrs", "name"]);
                            if (sAdj_name == "adj1") {
                                sAdj1 = PPTXXmlUtils.getTextByPathList(shapAdjst_ary[i], ["attrs", "fmla"]);
                                adj1 = parseInt(sAdj1.substr(4)) * SLIDE_FACTOR$1;
                            }
                            else if (sAdj_name == "adj2") {
                                sAdj2 = PPTXXmlUtils.getTextByPathList(shapAdjst_ary[i], ["attrs", "fmla"]);
                                adj2 = (parseInt(sAdj2.substr(4)) / 60000) * Math.PI / 180;
                            }
                            else if (sAdj_name == "adj3") {
                                sAdj3 = PPTXXmlUtils.getTextByPathList(shapAdjst_ary[i], ["attrs", "fmla"]);
                                adj3 = (parseInt(sAdj3.substr(4)) / 60000) * Math.PI / 180;
                            }
                            else if (sAdj_name == "adj4") {
                                sAdj4 = PPTXXmlUtils.getTextByPathList(shapAdjst_ary[i], ["attrs", "fmla"]);
                                adj4 = (parseInt(sAdj4.substr(4)) / 60000) * Math.PI / 180;
                            }
                            else if (sAdj_name == "adj5") {
                                sAdj5 = PPTXXmlUtils.getTextByPathList(shapAdjst_ary[i], ["attrs", "fmla"]);
                                adj5 = parseInt(sAdj5.substr(4)) * SLIDE_FACTOR$1;
                            }
                        }
                    }
                    var vc = h / 2, hc = w / 2, r = w, b = h, l = 0, t = 0, wd2 = w / 2, hd2 = h / 2;
                    const ss = Math.min(w, h);
                    var a5, maxAdj1, a1, enAng, stAng, th, thh, th2, rw1, rh1, rw2, rh2, rw3, rh3, wtH, htH, dxH, dyH, xH, yH, rI, u1, u2, u3, u4, u5, u6, u7, u8, u9, u10, u11, u12, u13, u14, u15, u16, u17, u18, u19, u20, u21, maxAng, aAng, ptAng, wtA, htA, dxA, dyA, xA, yA, wtE, htE, dxE, dyE, xE, yE, dxG, dyG, xG, yG, dxB, dyB, xB, yB, sx1, sy1, sx2, sy2, rO, x1O, y1O, x2O, y2O, dxO, dyO, dO, q1, q2, DO, q3, q4, q5, q6, q7, q8, sdelO, ndyO, sdyO, q9, q10, q11, dxF1, q12, dxF2, adyO, q13, q14, dyF1, q15, dyF2, q16, q17, q18, q19, q20, q21, q22, dxF, dyF, sdxF, sdyF, xF, yF, x1I, y1I, x2I, y2I, dxI, dyI, dI, v1, v2, DI, v3, v4, v5, v6, v7, v8, sdelI, v9, v10, v11, dxC1, v12, dxC2, adyI, v13, v14, dyC1, v15, dyC2, v16, v17, v18, v19, v20, v21, v22, dxC, dyC, sdxC, sdyC, xC, yC, ist0, ist1, istAng, isw1, isw2, iswAng, p1, p2, p3, p4, p5, xGp, yGp, xBp, yBp, en0, en1, en2, sw0, sw1, swAng;
                    const cnstVal1 = 25000 * SLIDE_FACTOR$1;
                    const cnstVal2 = 100000 * SLIDE_FACTOR$1;
                    const rdAngVal1 = (1 / 60000) * Math.PI / 180;
                    const rdAngVal2 = (21599999 / 60000) * Math.PI / 180;
                    const rdAngVal3 = 2 * Math.PI;
                    a5 = (adj5 < 0) ? 0 : (adj5 > cnstVal1) ? cnstVal1 : adj5;
                    maxAdj1 = a5 * 2;
                    a1 = (adj1 < 0) ? 0 : (adj1 > maxAdj1) ? maxAdj1 : adj1;
                    enAng = (adj3 < rdAngVal1) ? rdAngVal1 : (adj3 > rdAngVal2) ? rdAngVal2 : adj3;
                    stAng = (adj4 < 0) ? 0 : (adj4 > rdAngVal2) ? rdAngVal2 : adj4;
                    th = ss * a1 / cnstVal2;
                    thh = ss * a5 / cnstVal2;
                    th2 = th / 2;
                    rw1 = wd2 + th2 - thh;
                    rh1 = hd2 + th2 - thh;
                    rw2 = rw1 - th;
                    rh2 = rh1 - th;
                    rw3 = rw2 + th2;
                    rh3 = rh2 + th2;
                    wtH = rw3 * Math.sin(enAng);
                    htH = rh3 * Math.cos(enAng);
                    dxH = rw3 * Math.cos(Math.atan2(wtH, htH));
                    dyH = rh3 * Math.sin(Math.atan2(wtH, htH));
                    xH = hc + dxH;
                    yH = vc + dyH;
                    rI = (rw2 < rh2) ? rw2 : rh2;
                    u1 = dxH * dxH;
                    u2 = dyH * dyH;
                    u3 = rI * rI;
                    u4 = u1 - u3;
                    u5 = u2 - u3;
                    u6 = u4 * u5 / u1;
                    u7 = u6 / u2;
                    u8 = 1 - u7;
                    u9 = Math.sqrt(u8);
                    u10 = u4 / dxH;
                    u11 = u10 / dyH;
                    u12 = (1 + u9) / u11;
                    u13 = Math.atan2(u12, 1);
                    u14 = u13 + rdAngVal3;
                    u15 = (u13 > 0) ? u13 : u14;
                    u16 = u15 - enAng;
                    u17 = u16 + rdAngVal3;
                    u18 = (u16 > 0) ? u16 : u17;
                    u19 = u18 - cd2;
                    u20 = u18 - rdAngVal3;
                    u21 = (u19 > 0) ? u20 : u18;
                    maxAng = Math.abs(u21);
                    aAng = (adj2 < 0) ? 0 : (adj2 > maxAng) ? maxAng : adj2;
                    ptAng = enAng + aAng;
                    wtA = rw3 * Math.sin(ptAng);
                    htA = rh3 * Math.cos(ptAng);
                    dxA = rw3 * Math.cos(Math.atan2(wtA, htA));
                    dyA = rh3 * Math.sin(Math.atan2(wtA, htA));
                    xA = hc + dxA;
                    yA = vc + dyA;
                    wtE = rw1 * Math.sin(stAng);
                    htE = rh1 * Math.cos(stAng);
                    dxE = rw1 * Math.cos(Math.atan2(wtE, htE));
                    dyE = rh1 * Math.sin(Math.atan2(wtE, htE));
                    xE = hc + dxE;
                    yE = vc + dyE;
                    dxG = thh * Math.cos(ptAng);
                    dyG = thh * Math.sin(ptAng);
                    xG = xH + dxG;
                    yG = yH + dyG;
                    dxB = thh * Math.cos(ptAng);
                    dyB = thh * Math.sin(ptAng);
                    xB = xH - dxB;
                    yB = yH - dyB;
                    sx1 = xB - hc;
                    sy1 = yB - vc;
                    sx2 = xG - hc;
                    sy2 = yG - vc;
                    rO = (rw1 < rh1) ? rw1 : rh1;
                    x1O = sx1 * rO / rw1;
                    y1O = sy1 * rO / rh1;
                    x2O = sx2 * rO / rw1;
                    y2O = sy2 * rO / rh1;
                    dxO = x2O - x1O;
                    dyO = y2O - y1O;
                    dO = Math.sqrt(dxO * dxO + dyO * dyO);
                    q1 = x1O * y2O;
                    q2 = x2O * y1O;
                    DO = q1 - q2;
                    q3 = rO * rO;
                    q4 = dO * dO;
                    q5 = q3 * q4;
                    q6 = DO * DO;
                    q7 = q5 - q6;
                    q8 = (q7 > 0) ? q7 : 0;
                    sdelO = Math.sqrt(q8);
                    ndyO = dyO * -1;
                    sdyO = (ndyO > 0) ? -1 : 1;
                    q9 = sdyO * dxO;
                    q10 = q9 * sdelO;
                    q11 = DO * dyO;
                    dxF1 = (q11 + q10) / q4;
                    q12 = q11 - q10;
                    dxF2 = q12 / q4;
                    adyO = Math.abs(dyO);
                    q13 = adyO * sdelO;
                    q14 = DO * dxO / -1;
                    dyF1 = (q14 + q13) / q4;
                    q15 = q14 - q13;
                    dyF2 = q15 / q4;
                    q16 = x2O - dxF1;
                    q17 = x2O - dxF2;
                    q18 = y2O - dyF1;
                    q19 = y2O - dyF2;
                    q20 = Math.sqrt(q16 * q16 + q18 * q18);
                    q21 = Math.sqrt(q17 * q17 + q19 * q19);
                    q22 = q21 - q20;
                    dxF = (q22 > 0) ? dxF1 : dxF2;
                    dyF = (q22 > 0) ? dyF1 : dyF2;
                    sdxF = dxF * rw1 / rO;
                    sdyF = dyF * rh1 / rO;
                    xF = hc + sdxF;
                    yF = vc + sdyF;
                    x1I = sx1 * rI / rw2;
                    y1I = sy1 * rI / rh2;
                    x2I = sx2 * rI / rw2;
                    y2I = sy2 * rI / rh2;
                    dxI = x2I - x1I;
                    dyI = y2I - y1I;
                    dI = Math.sqrt(dxI * dxI + dyI * dyI);
                    v1 = x1I * y2I;
                    v2 = x2I * y1I;
                    DI = v1 - v2;
                    v3 = rI * rI;
                    v4 = dI * dI;
                    v5 = v3 * v4;
                    v6 = DI * DI;
                    v7 = v5 - v6;
                    v8 = (v7 > 0) ? v7 : 0;
                    sdelI = Math.sqrt(v8);
                    v9 = sdyO * dxI;
                    v10 = v9 * sdelI;
                    v11 = DI * dyI;
                    dxC1 = (v11 + v10) / v4;
                    v12 = v11 - v10;
                    dxC2 = v12 / v4;
                    adyI = Math.abs(dyI);
                    v13 = adyI * sdelI;
                    v14 = DI * dxI / -1;
                    dyC1 = (v14 + v13) / v4;
                    v15 = v14 - v13;
                    dyC2 = v15 / v4;
                    v16 = x1I - dxC1;
                    v17 = x1I - dxC2;
                    v18 = y1I - dyC1;
                    v19 = y1I - dyC2;
                    v20 = Math.sqrt(v16 * v16 + v18 * v18);
                    v21 = Math.sqrt(v17 * v17 + v19 * v19);
                    v22 = v21 - v20;
                    dxC = (v22 > 0) ? dxC1 : dxC2;
                    dyC = (v22 > 0) ? dyC1 : dyC2;
                    sdxC = dxC * rw2 / rI;
                    sdyC = dyC * rh2 / rI;
                    xC = hc + sdxC;
                    yC = vc + sdyC;
                    ist0 = Math.atan2(sdyC, sdxC);
                    ist1 = ist0 + rdAngVal3;
                    istAng = (ist0 > 0) ? ist0 : ist1;
                    isw1 = stAng - istAng;
                    isw2 = isw1 - rdAngVal3;
                    iswAng = (isw1 > 0) ? isw2 : isw1;
                    p1 = xF - xC;
                    p2 = yF - yC;
                    p3 = Math.sqrt(p1 * p1 + p2 * p2);
                    p4 = p3 / 2;
                    p5 = p4 - thh;
                    xGp = (p5 > 0) ? xF : xG;
                    yGp = (p5 > 0) ? yF : yG;
                    xBp = (p5 > 0) ? xC : xB;
                    yBp = (p5 > 0) ? yC : yB;
                    en0 = Math.atan2(sdyF, sdxF);
                    en1 = en0 + rdAngVal3;
                    en2 = (en0 > 0) ? en0 : en1;
                    sw0 = en2 - stAng;
                    sw1 = sw0 + rdAngVal3;
                    swAng = (sw0 > 0) ? sw0 : sw1;
                    const strtAng = stAng * 180 / Math.PI;
                    const endAng = strtAng + (swAng * 180 / Math.PI);
                    const stiAng = istAng * 180 / Math.PI;
                    const swiAng = iswAng * 180 / Math.PI;
                    const ediAng = stiAng + swiAng;
                    let d_val = `${PPTXShapeUtils.shapeArc(w / 2, h / 2, rw1, rh1, strtAng, endAng, false)} L${xGp},${yGp} L${xA},${yA} L${xBp},${yBp} L${xC},${yC}${PPTXShapeUtils.shapeArc(w / 2, h / 2, rw2, rh2, stiAng, ediAng, false).replace("M", "L")} z`;
                    result += `<path d='${d_val}' fill='${(!imgFillFlg ? (grndFillFlg ? "url(#linGrd_" + shpId + ")" : fillColor) : "url(#imgPtrn_" + shpId + ")")}' stroke='${border.color}' stroke-width='${border.width}' stroke-dasharray='${border.strokeDasharray}' />`;
                    break;
                }
                case "leftCircularArrow": {
                    const shapAdjst_ary = PPTXXmlUtils.getTextByPathList(node, ["p:spPr", "a:prstGeom", "a:avLst", "a:gd"]);
                    let sAdj1, adj1 = 12500 * SLIDE_FACTOR$1;
                    let sAdj2, adj2 = (-1142319 / 60000) * Math.PI / 180;
                    let sAdj3, adj3 = (1142319 / 60000) * Math.PI / 180;
                    let sAdj4, adj4 = (10800000 / 60000) * Math.PI / 180;
                    let sAdj5, adj5 = 12500 * SLIDE_FACTOR$1;
                    if (shapAdjst_ary !== undefined) {
                        for (const i of shapAdjst_ary.keys()) {
                            const sAdj_name = PPTXXmlUtils.getTextByPathList(shapAdjst_ary[i], ["attrs", "name"]);
                            if (sAdj_name == "adj1") {
                                sAdj1 = PPTXXmlUtils.getTextByPathList(shapAdjst_ary[i], ["attrs", "fmla"]);
                                adj1 = parseInt(sAdj1.substr(4)) * SLIDE_FACTOR$1;
                            }
                            else if (sAdj_name == "adj2") {
                                sAdj2 = PPTXXmlUtils.getTextByPathList(shapAdjst_ary[i], ["attrs", "fmla"]);
                                adj2 = (parseInt(sAdj2.substr(4)) / 60000) * Math.PI / 180;
                            }
                            else if (sAdj_name == "adj3") {
                                sAdj3 = PPTXXmlUtils.getTextByPathList(shapAdjst_ary[i], ["attrs", "fmla"]);
                                adj3 = (parseInt(sAdj3.substr(4)) / 60000) * Math.PI / 180;
                            }
                            else if (sAdj_name == "adj4") {
                                sAdj4 = PPTXXmlUtils.getTextByPathList(shapAdjst_ary[i], ["attrs", "fmla"]);
                                adj4 = (parseInt(sAdj4.substr(4)) / 60000) * Math.PI / 180;
                            }
                            else if (sAdj_name == "adj5") {
                                sAdj5 = PPTXXmlUtils.getTextByPathList(shapAdjst_ary[i], ["attrs", "fmla"]);
                                adj5 = parseInt(sAdj5.substr(4)) * SLIDE_FACTOR$1;
                            }
                        }
                    }
                    var vc = h / 2, hc = w / 2, r = w, b = h, l = 0, t = 0, wd2 = w / 2, hd2 = h / 2;
                    const ss = Math.min(w, h);
                    const cnstVal1 = 25000 * SLIDE_FACTOR$1;
                    const cnstVal2 = 100000 * SLIDE_FACTOR$1;
                    const rdAngVal1 = (1 / 60000) * Math.PI / 180;
                    const rdAngVal2 = (21599999 / 60000) * Math.PI / 180;
                    const rdAngVal3 = 2 * Math.PI;
                    var a5, maxAdj1, a1, enAng, stAng, th, thh, th2, rw1, rh1, rw2, rh2, rw3, rh3, wtH, htH, dxH, dyH, xH, yH, rI, u1, u2, u3, u4, u5, u6, u7, u8, u9, u10, u11, u12, u13, u14, u15, u16, u17, u18, u19, u20, u21, u22, minAng, u23, a2, aAng, ptAng, wtA, htA, dxA, dyA, xA, yA, wtE, htE, dxE, dyE, xE, yE, wtD, htD, dxD, dyD, xD, yD, dxG, dyG, xG, yG, dxB, dyB, xB, yB, sx1, sy1, sx2, sy2, rO, x1O, y1O, x2O, y2O, dxO, dyO, dO, q1, q2, DO, q3, q4, q5, q6, q7, q8, sdelO, ndyO, sdyO, q9, q10, q11, dxF1, q12, dxF2, adyO, q13, q14, dyF1, q15, dyF2, q16, q17, q18, q19, q20, q21, q22, dxF, dyF, sdxF, sdyF, xF, yF, x1I, y1I, x2I, y2I, dxI, dyI, dI, v1, v2, DI, v3, v4, v5, v6, v7, v8, sdelI, v9, v10, v11, dxC1, v12, dxC2, adyI, v13, v14, dyC1, v15, dyC2, v16, v17, v18, v19, v20, v21, v22, dxC, dyC, sdxC, sdyC, xC, yC, ist0, ist1, istAng0, isw1, isw2, iswAng0, istAng, iswAng, p1, p2, p3, p4, p5, xGp, yGp, xBp, yBp, en0, en1, en2, sw0, sw1, swAng, stAng0;
                    a5 = (adj5 < 0) ? 0 : (adj5 > cnstVal1) ? cnstVal1 : adj5;
                    maxAdj1 = a5 * 2;
                    a1 = (adj1 < 0) ? 0 : (adj1 > maxAdj1) ? maxAdj1 : adj1;
                    enAng = (adj3 < rdAngVal1) ? rdAngVal1 : (adj3 > rdAngVal2) ? rdAngVal2 : adj3;
                    stAng = (adj4 < 0) ? 0 : (adj4 > rdAngVal2) ? rdAngVal2 : adj4;
                    th = ss * a1 / cnstVal2;
                    thh = ss * a5 / cnstVal2;
                    th2 = th / 2;
                    rw1 = wd2 + th2 - thh;
                    rh1 = hd2 + th2 - thh;
                    rw2 = rw1 - th;
                    rh2 = rh1 - th;
                    rw3 = rw2 + th2;
                    rh3 = rh2 + th2;
                    wtH = rw3 * Math.sin(enAng);
                    htH = rh3 * Math.cos(enAng);
                    dxH = rw3 * Math.cos(Math.atan2(wtH, htH));
                    dyH = rh3 * Math.sin(Math.atan2(wtH, htH));
                    xH = hc + dxH;
                    yH = vc + dyH;
                    rI = (rw2 < rh2) ? rw2 : rh2;
                    u1 = dxH * dxH;
                    u2 = dyH * dyH;
                    u3 = rI * rI;
                    u4 = u1 - u3;
                    u5 = u2 - u3;
                    u6 = u4 * u5 / u1;
                    u7 = u6 / u2;
                    u8 = 1 - u7;
                    u9 = Math.sqrt(u8);
                    u10 = u4 / dxH;
                    u11 = u10 / dyH;
                    u12 = (1 + u9) / u11;
                    u13 = Math.atan2(u12, 1);
                    u14 = u13 + rdAngVal3;
                    u15 = (u13 > 0) ? u13 : u14;
                    u16 = u15 - enAng;
                    u17 = u16 + rdAngVal3;
                    u18 = (u16 > 0) ? u16 : u17;
                    u19 = u18 - cd2;
                    u20 = u18 - rdAngVal3;
                    u21 = (u19 > 0) ? u20 : u18;
                    u22 = Math.abs(u21);
                    minAng = u22 * -1;
                    u23 = Math.abs(adj2);
                    a2 = u23 * -1;
                    aAng = (a2 < minAng) ? minAng : (a2 > 0) ? 0 : a2;
                    ptAng = enAng + aAng;
                    wtA = rw3 * Math.sin(ptAng);
                    htA = rh3 * Math.cos(ptAng);
                    dxA = rw3 * Math.cos(Math.atan2(wtA, htA));
                    dyA = rh3 * Math.sin(Math.atan2(wtA, htA));
                    xA = hc + dxA;
                    yA = vc + dyA;
                    wtE = rw1 * Math.sin(stAng);
                    htE = rh1 * Math.cos(stAng);
                    dxE = rw1 * Math.cos(Math.atan2(wtE, htE));
                    dyE = rh1 * Math.sin(Math.atan2(wtE, htE));
                    xE = hc + dxE;
                    yE = vc + dyE;
                    wtD = rw2 * Math.sin(stAng);
                    htD = rh2 * Math.cos(stAng);
                    dxD = rw2 * Math.cos(Math.atan2(wtD, htD));
                    dyD = rh2 * Math.sin(Math.atan2(wtD, htD));
                    xD = hc + dxD;
                    yD = vc + dyD;
                    dxG = thh * Math.cos(ptAng);
                    dyG = thh * Math.sin(ptAng);
                    xG = xH + dxG;
                    yG = yH + dyG;
                    dxB = thh * Math.cos(ptAng);
                    dyB = thh * Math.sin(ptAng);
                    xB = xH - dxB;
                    yB = yH - dyB;
                    sx1 = xB - hc;
                    sy1 = yB - vc;
                    sx2 = xG - hc;
                    sy2 = yG - vc;
                    rO = (rw1 < rh1) ? rw1 : rh1;
                    x1O = sx1 * rO / rw1;
                    y1O = sy1 * rO / rh1;
                    x2O = sx2 * rO / rw1;
                    y2O = sy2 * rO / rh1;
                    dxO = x2O - x1O;
                    dyO = y2O - y1O;
                    dO = Math.sqrt(dxO * dxO + dyO * dyO);
                    q1 = x1O * y2O;
                    q2 = x2O * y1O;
                    DO = q1 - q2;
                    q3 = rO * rO;
                    q4 = dO * dO;
                    q5 = q3 * q4;
                    q6 = DO * DO;
                    q7 = q5 - q6;
                    q8 = (q7 > 0) ? q7 : 0;
                    sdelO = Math.sqrt(q8);
                    ndyO = dyO * -1;
                    sdyO = (ndyO > 0) ? -1 : 1;
                    q9 = sdyO * dxO;
                    q10 = q9 * sdelO;
                    q11 = DO * dyO;
                    dxF1 = (q11 + q10) / q4;
                    q12 = q11 - q10;
                    dxF2 = q12 / q4;
                    adyO = Math.abs(dyO);
                    q13 = adyO * sdelO;
                    q14 = DO * dxO / -1;
                    dyF1 = (q14 + q13) / q4;
                    q15 = q14 - q13;
                    dyF2 = q15 / q4;
                    q16 = x2O - dxF1;
                    q17 = x2O - dxF2;
                    q18 = y2O - dyF1;
                    q19 = y2O - dyF2;
                    q20 = Math.sqrt(q16 * q16 + q18 * q18);
                    q21 = Math.sqrt(q17 * q17 + q19 * q19);
                    q22 = q21 - q20;
                    dxF = (q22 > 0) ? dxF1 : dxF2;
                    dyF = (q22 > 0) ? dyF1 : dyF2;
                    sdxF = dxF * rw1 / rO;
                    sdyF = dyF * rh1 / rO;
                    xF = hc + sdxF;
                    yF = vc + sdyF;
                    x1I = sx1 * rI / rw2;
                    y1I = sy1 * rI / rh2;
                    x2I = sx2 * rI / rw2;
                    y2I = sy2 * rI / rh2;
                    dxI = x2I - x1I;
                    dyI = y2I - y1I;
                    dI = Math.sqrt(dxI * dxI + dyI * dyI);
                    v1 = x1I * y2I;
                    v2 = x2I * y1I;
                    DI = v1 - v2;
                    v3 = rI * rI;
                    v4 = dI * dI;
                    v5 = v3 * v4;
                    v6 = DI * DI;
                    v7 = v5 - v6;
                    v8 = (v7 > 0) ? v7 : 0;
                    sdelI = Math.sqrt(v8);
                    v9 = sdyO * dxI;
                    v10 = v9 * sdelI;
                    v11 = DI * dyI;
                    dxC1 = (v11 + v10) / v4;
                    v12 = v11 - v10;
                    dxC2 = v12 / v4;
                    adyI = Math.abs(dyI);
                    v13 = adyI * sdelI;
                    v14 = DI * dxI / -1;
                    dyC1 = (v14 + v13) / v4;
                    v15 = v14 - v13;
                    dyC2 = v15 / v4;
                    v16 = x1I - dxC1;
                    v17 = x1I - dxC2;
                    v18 = y1I - dyC1;
                    v19 = y1I - dyC2;
                    v20 = Math.sqrt(v16 * v16 + v18 * v18);
                    v21 = Math.sqrt(v17 * v17 + v19 * v19);
                    v22 = v21 - v20;
                    dxC = (v22 > 0) ? dxC1 : dxC2;
                    dyC = (v22 > 0) ? dyC1 : dyC2;
                    sdxC = dxC * rw2 / rI;
                    sdyC = dyC * rh2 / rI;
                    xC = hc + sdxC;
                    yC = vc + sdyC;
                    ist0 = Math.atan2(sdyC, sdxC);
                    ist1 = ist0 + rdAngVal3;
                    istAng0 = (ist0 > 0) ? ist0 : ist1;
                    isw1 = stAng - istAng0;
                    isw2 = isw1 + rdAngVal3;
                    iswAng0 = (isw1 > 0) ? isw1 : isw2;
                    istAng = istAng0 + iswAng0;
                    iswAng = -iswAng0;
                    p1 = xF - xC;
                    p2 = yF - yC;
                    p3 = Math.sqrt(p1 * p1 + p2 * p2);
                    p4 = p3 / 2;
                    p5 = p4 - thh;
                    xGp = (p5 > 0) ? xF : xG;
                    yGp = (p5 > 0) ? yF : yG;
                    xBp = (p5 > 0) ? xC : xB;
                    yBp = (p5 > 0) ? yC : yB;
                    en0 = Math.atan2(sdyF, sdxF);
                    en1 = en0 + rdAngVal3;
                    en2 = (en0 > 0) ? en0 : en1;
                    sw0 = en2 - stAng;
                    sw1 = sw0 - rdAngVal3;
                    swAng = (sw0 > 0) ? sw1 : sw0;
                    stAng0 = stAng + swAng;
                    const strtAng = stAng0 * 180 / Math.PI;
                    const endAng = stAng * 180 / Math.PI;
                    const stiAng = istAng * 180 / Math.PI;
                    const swiAng = iswAng * 180 / Math.PI;
                    const ediAng = stiAng + swiAng;
                    let d_val = `M${xE},${yE} L${xD},${yD}${PPTXShapeUtils.shapeArc(w / 2, h / 2, rw2, rh2, stiAng, ediAng, false).replace("M", "L")} L${xBp},${yBp} L${xA},${yA} L${xGp},${yGp} L${xF},${yF}${PPTXShapeUtils.shapeArc(w / 2, h / 2, rw1, rh1, strtAng, endAng, false).replace("M", "L")} z`;
                    result += `<path d='${d_val}' fill='${(!imgFillFlg ? (grndFillFlg ? "url(#linGrd_" + shpId + ")" : fillColor) : "url(#imgPtrn_" + shpId + ")")}' stroke='${border.color}' stroke-width='${border.width}' stroke-dasharray='${border.strokeDasharray}' />`;
                    break;
                }
                case "funnel": {
                    const shapAdjst = PPTXXmlUtils.getTextByPathList(node, ["p:spPr", "a:prstGeom", "a:avLst", "a:gd", "attrs", "fmla"]);
                    let adj = 40000 * SLIDE_FACTOR$1;
                    if (shapAdjst !== undefined) {
                        adj = parseInt(shapAdjst.substr(4)) * SLIDE_FACTOR$1;
                    }
                    const cnstVal2 = 100000 * SLIDE_FACTOR$1;
                    let a = (adj < 0) ? 0 : (adj > cnstVal2) ? cnstVal2 : adj;
                    const bottomW = w * a / cnstVal2;
                    var d = `M0,0 L${w},0 L${((w + bottomW) / 2)},${h} L${((w - bottomW) / 2)},${h} z`;
                    result += `<path d='${d}' fill='${(!imgFillFlg ? (grndFillFlg ? "url(#linGrd_" + shpId + ")" : fillColor) : "url(#imgPtrn_" + shpId + ")")}' stroke='${border.color}' stroke-width='${border.width}' stroke-dasharray='${border.strokeDasharray}' />`;
                    break;
                }
                case "leftRightCircularArrow": {
                    var wd2 = w / 2;
                    let hd2 = h / 2;
                    let r = Math.min(wd2, hd2);
                    var d = `M${(wd2 - r)},${hd2}${PPTXShapeUtils.shapeArc(wd2, hd2, r, r, 180, 360, false).replace("M", "L")} M${(wd2 - r - r * 0.3)},${(hd2 - r * 0.2)} L${(wd2 - r)},${hd2} L${(wd2 - r - r * 0.3)},${(hd2 + r * 0.2)} M${(wd2 + r + r * 0.3)},${(hd2 - r * 0.2)} L${(wd2 + r)},${hd2} L${(wd2 + r + r * 0.3)},${(hd2 + r * 0.2)}`;
                    result += `<path d='${d}' fill='${(!imgFillFlg ? (grndFillFlg ? "url(#linGrd_" + shpId + ")" : fillColor) : "url(#imgPtrn_" + shpId + ")")}' stroke='${border.color}' stroke-width='${border.width}' stroke-dasharray='${border.strokeDasharray}' />`;
                    break;
                }
                case "flowChartOfflineStorage": {
                    var d = `M0,0 L${w},0 L${w},${(h * 0.7)} L${(w * 0.66)},${h} L${(w * 0.34)},${h} L0,${(h * 0.7)} z`;
                    result += `<path d='${d}' fill='${(!imgFillFlg ? (grndFillFlg ? "url(#linGrd_" + shpId + ")" : fillColor) : "url(#imgPtrn_" + shpId + ")")}' stroke='${border.color}' stroke-width='${border.width}' stroke-dasharray='${border.strokeDasharray}' />`;
                    break;
                }
            }
            result += "</svg>";
            const dataAttrs1 = genShapeDataAttributes(node, workingXfrmNode, id, name, idx, type, rotate, sType);
            const animationData = extractAnimationData(node, warpObj);
            let animationAttrs = "";
            if (animationData) {
                animationAttrs = ` data-animation='${JSON.stringify(animationData)}'`;
            }
            result += `<div class='block ${PPTXStyleUtils.getVerticalAlign(node, slideLayoutSpNode, slideMasterSpNode, type)} ${PPTXStyleUtils.getContentDir(node, type, warpObj)}' _id='${id}' _idx='${idx}' _type='${type}' _name='${name}' style='${PPTXXmlUtils.getPosition(workingXfrmNode, pNode, slideLayoutXfrmNode, slideMasterXfrmNode, sType)}${PPTXXmlUtils.getSize(workingXfrmNode, slideLayoutXfrmNode, slideMasterXfrmNode)}${transform3dStyle} z-index: ${order};'${dataAttrs1}${animationAttrs}>`;
            if (node["p:txBody"] !== undefined && (isUserDrawnBg === undefined || isUserDrawnBg === true)) {
                if (type != "diagram" && type != "textBox") {
                    type = "shape";
                }
                result += await PPTXTextUtils.genTextBody(node["p:txBody"], node, slideLayoutSpNode, slideMasterSpNode, type, idx, warpObj);
            }
            result += "</div>";
        }
        else if (custShapType !== undefined) {
            const renderW = (sType === 'group-abs') ? w : drawW;
            const renderH = (sType === 'group-abs') ? h : drawH;
            result += renderCustomShape(custShapType, renderW, renderH, imgFillFlg, grndFillFlg, fillColor, border, shpId, shapeArc);
            result += "</svg>";
            const dataAttrs2 = genShapeDataAttributes(node, workingXfrmNode, id, name, idx, type, rotate, sType);
            const animationData2 = extractAnimationData(node, warpObj);
            let animationAttrs2 = "";
            if (animationData2) {
                animationAttrs2 = ` data-animation='${JSON.stringify(animationData2)}'`;
            }
            result += `<div class='block ${PPTXStyleUtils.getVerticalAlign(node, slideLayoutSpNode, slideMasterSpNode, type)} ${PPTXStyleUtils.getContentDir(node, type, warpObj)}' _id='${id}' _idx='${idx}' _type='${type}' _name='${name}' style='${PPTXXmlUtils.getPosition(workingXfrmNode, pNode, slideLayoutXfrmNode, slideMasterXfrmNode, sType)}${PPTXXmlUtils.getSize(workingXfrmNode, slideLayoutXfrmNode, slideMasterXfrmNode)} z-index: ${order};'${dataAttrs2}${animationAttrs2}>`;
            if (node["p:txBody"] !== undefined && (isUserDrawnBg === undefined || isUserDrawnBg === true)) {
                if (type != "diagram" && type != "textBox") {
                    type = "shape";
                }
                let textNode = node;
                if (sType === 'group-abs' && workingXfrmNode !== slideXfrmNode) {
                    textNode = JSON.parse(JSON.stringify(node));
                    if (textNode["p:spPr"] && textNode["p:spPr"]["a:xfrm"]) {
                        textNode["p:spPr"]["a:xfrm"] = workingXfrmNode;
                    }
                }
                result += await PPTXTextUtils.genTextBody(textNode["p:txBody"], textNode, slideLayoutSpNode, slideMasterSpNode, type, idx, warpObj);
            }
            result += "</div>";
        }
        else {
            const dataAttrs3 = genShapeDataAttributes(node, slideXfrmNode, id, name, idx, type, rotate, sType);
            const animationData3 = extractAnimationData(node, warpObj);
            if (animationData3) {
                ` data-animation='${JSON.stringify(animationData3)}'`;
            }
            result += `<div class='block ${PPTXStyleUtils.getVerticalAlign(node, slideLayoutSpNode, slideMasterSpNode, type)} ${PPTXStyleUtils.getContentDir(node, type, warpObj)}' _id='${id}' _idx='${idx}' _type='${type}' _name='${name}' style='${PPTXXmlUtils.getPosition(slideXfrmNode, pNode, slideLayoutXfrmNode, slideMasterXfrmNode, sType)}${PPTXXmlUtils.getSize(slideXfrmNode, slideLayoutXfrmNode, slideMasterXfrmNode)}${PPTXStyleUtils.getBorder(node, pNode, false, "shape", warpObj)}${await PPTXStyleUtils.getShapeFill(node, pNode, false, warpObj, source)} z-index: ${order};'${dataAttrs3}>`;
            if (node["p:txBody"] !== undefined && (isUserDrawnBg === undefined || isUserDrawnBg === true)) {
                result += await PPTXTextUtils.genTextBody(node["p:txBody"], node, slideLayoutSpNode, slideMasterSpNode, type, idx, warpObj);
            }
            result += "</div>";
        }
        return result;
    }
    return {
        shapeArc: shapeArc,
        shapeArcAlt: shapeArcAlt,
        shapePie: shapePie,
        shapeGear: shapeGear,
        shapeSnipRoundRect: shapeSnipRoundRect,
        shapeSnipRoundRectAlt: shapeSnipRoundRectAlt,
        polarToCartesian: polarToCartesian,
        genShape,
    };
    function extractAnimationData(node, warpObj) {
        const nvSpPr = PPTXXmlUtils.getTextByPathList(node, ["p:nvSpPr"]);
        if (!nvSpPr)
            return null;
        const nvPr = PPTXXmlUtils.getTextByPathList(nvSpPr, ["p:nvPr"]);
        if (!nvPr)
            return null;
        const animLst = PPTXXmlUtils.getTextByPathList(nvPr, ["p:animLst"]);
        if (animLst) {
            return parseAnimationList(animLst);
        }
        const animRef = PPTXXmlUtils.getTextByPathList(nvPr, ["p:animRef"]);
        if (animRef && animRef.attrs) {
            const rId = animRef.attrs["r:embed"];
            if (rId && warpObj.slideAnims && warpObj.slideAnims[rId]) {
                return warpObj.slideAnims[rId];
            }
        }
        return null;
    }
    function parseAnimationList(animLst) {
        const animArray = Array.isArray(animLst["p:par"]) ? animLst["p:par"] :
            (animLst["p:par"] ? [animLst["p:par"]] : []);
        if (animArray.length === 0)
            return null;
        const par = animArray[0];
        if (par["p:cTn"]) {
            const cTn = par["p:cTn"];
            const animType = getAnimationType(cTn);
            const duration = cTn.attrs["dur"] || "1000";
            const delay = cTn.attrs["st"] || "0";
            return {
                type: animType,
                duration: parseInt(duration),
                delay: parseInt(delay)
            };
        }
        return null;
    }
    function getAnimationType(cTn) {
        if (cTn["p:childTnLst"]) {
            const childTnLst = cTn["p:childTnLst"];
            if (childTnLst["p:set"]) {
                const set = childTnLst["p:set"];
                if (set["p:to"]) {
                    const to = set["p:to"];
                    if (to["p:strVal"] && to["p:strVal"].attrs["val"] === "visible") {
                        return "fade-in";
                    }
                }
            }
            if (childTnLst["p:cmd"]) {
                return "custom";
            }
        }
        return "appear";
    }
    function process3DEffects(scene3d, sp3d) {
        let transform = "";
        if (scene3d && scene3d["a:camera"]) {
            const camera = scene3d["a:camera"];
            const prst = camera.attrs?.["prst"];
            switch (prst) {
                case "orthographicFront":
                    break;
                case "orthographicTop":
                    transform += " rotateX(-90deg)";
                    break;
                case "orthographicBottom":
                    transform += " rotateX(90deg)";
                    break;
                case "orthographicLeft":
                    transform += " rotateY(90deg)";
                    break;
                case "orthographicRight":
                    transform += " rotateY(-90deg)";
                    break;
                case "perspectiveFront":
                    transform += " perspective(1000px)";
                    break;
                case "perspectiveTop":
                    transform += " perspective(1000px) rotateX(-60deg)";
                    break;
                case "perspectiveBottom":
                    transform += " perspective(1000px) rotateX(60deg)";
                    break;
                case "perspectiveLeft":
                    transform += " perspective(1000px) rotateY(60deg)";
                    break;
                case "perspectiveRight":
                    transform += " perspective(1000px) rotateY(-60deg)";
                    break;
                default:
                    transform += " perspective(800px)";
            }
        }
        if (sp3d) {
            if (sp3d["a:extrusionH"]) {
                const extrusionH = parseInt(sp3d["a:extrusionH"].attrs?.["val"] || "0");
                if (extrusionH > 0) {
                    const depth = Math.round(extrusionH * SLIDE_FACTOR$1);
                    if (depth > 0) {
                        transform += ` translateZ(${depth}px)`;
                    }
                }
            }
            if (sp3d["a:bevelT"]) {
                const bevelT = sp3d["a:bevelT"];
                const w = parseInt(bevelT.attrs?.["w"] || "0");
                const h = parseInt(bevelT.attrs?.["h"] || "0");
                if (w > 0 || h > 0) {
                    transform += " rotateX(5deg) rotateY(5deg)";
                }
            }
            if (sp3d["a:bevelB"]) {
                const bevelB = sp3d["a:bevelB"];
                const w = parseInt(bevelB.attrs?.["w"] || "0");
                const h = parseInt(bevelB.attrs?.["h"] || "0");
                if (w > 0 || h > 0) {
                    transform += " rotateX(-3deg) rotateY(-3deg)";
                }
            }
        }
        if (transform) {
            return ` transform:${transform}; transform-style: preserve-3d;`;
        }
        return "";
    }
})();

async function genChart(node, warpObj, parentNode) {
    const order = node["attrs"]["order"];
    let xfrmNode = PPTXXmlUtils.getTextByPathList(node, ["a:xfrm"]) ||
        PPTXXmlUtils.getTextByPathList(node, ["p:xfrm"]);
    let workingXfrmNode = xfrmNode;
    if (warpObj.currentGroupScale && xfrmNode) {
        const { scaleX, scaleY, childX, childY } = warpObj.currentGroupScale;
        workingXfrmNode = JSON.parse(JSON.stringify(xfrmNode));
        if (xfrmNode['a:ext'] && xfrmNode['a:ext'].attrs) {
            const originalCx = parseInt(xfrmNode['a:ext'].attrs.cx);
            const originalCy = parseInt(xfrmNode['a:ext'].attrs.cy);
            workingXfrmNode['a:ext'].attrs.cx = Math.round(originalCx * scaleX);
            workingXfrmNode['a:ext'].attrs.cy = Math.round(originalCy * scaleY);
        }
        if (xfrmNode['a:off'] && xfrmNode['a:off'].attrs) {
            const originalOffX = parseInt(xfrmNode['a:off'].attrs.x);
            const originalOffY = parseInt(xfrmNode['a:off'].attrs.y);
            const relativeX = originalOffX - (childX / SLIDE_FACTOR$1);
            const relativeY = originalOffY - (childY / SLIDE_FACTOR$1);
            workingXfrmNode['a:off'].attrs.x = Math.round(childX / SLIDE_FACTOR$1 + relativeX * scaleX);
            workingXfrmNode['a:off'].attrs.y = Math.round(childY / SLIDE_FACTOR$1 + relativeY * scaleY);
        }
    }
    let offX = 0, offY = 0, extCx = 0, extCy = 0;
    if (workingXfrmNode !== undefined) {
        if (workingXfrmNode['a:off'] && workingXfrmNode['a:off'].attrs) {
            offX = workingXfrmNode['a:off'].attrs.x || 0;
            offY = workingXfrmNode['a:off'].attrs.y || 0;
        }
        if (workingXfrmNode['a:ext'] && workingXfrmNode['a:ext'].attrs) {
            extCx = workingXfrmNode['a:ext'].attrs.cx || 0;
            extCy = workingXfrmNode['a:ext'].attrs.cy || 0;
        }
    }
    const dataAttrs = ` data-node-type="chart" data-off-x="${offX}" data-off-y="${offY}" data-ext-cx="${extCx}" data-ext-cy="${extCy}"`;
    const result = `<div id='chart${warpObj.chartId.value}' class='block content' style='${PPTXXmlUtils.getPosition(workingXfrmNode, parentNode || node, undefined, undefined)}${PPTXXmlUtils.getSize(workingXfrmNode, undefined, undefined)}` +
        ` z-index: ${order};'${dataAttrs}></div>`;
    const rid = node["a:graphic"]["a:graphicData"]["c:chart"]["attrs"]["r:id"];
    const refName = warpObj["slideResObj"][rid]["target"];
    const content = await PPTXXmlUtils.readXmlFile(warpObj["zip"], refName);
    if (!content) {
        return result;
    }
    const chartSpace = PPTXXmlUtils.getTextByPathList(content, ["c:chartSpace"]);
    if (!chartSpace) {
        return result;
    }
    const chart = PPTXXmlUtils.getTextByPathList(chartSpace, ["c:chart"]);
    const plotArea = PPTXXmlUtils.getTextByPathList(chart, ["c:plotArea"]);
    const view3D = PPTXXmlUtils.getTextByPathList(chart, ["c:view3D"]);
    const view3DProps = {};
    if (view3D) {
        if (view3D["attrs"]?.rotX !== undefined)
            view3DProps.rotX = parseFloat(view3D["attrs"].rotX);
        if (view3D["attrs"]?.rotY !== undefined)
            view3DProps.rotY = parseFloat(view3D["attrs"].rotY);
        if (view3D["attrs"]?.depthPercent !== undefined)
            view3DProps.depthPercent = parseFloat(view3D["attrs"].depthPercent);
        if (view3D["attrs"]?.rAngAx !== undefined)
            view3DProps.rAngAx = view3D["attrs"].rAngAx === "1";
    }
    const chartType = Object.keys(plotArea).find(key => key.startsWith('c:') && key.endsWith('Chart'));
    const varyColors = chartType ? PPTXXmlUtils.getTextByPathList(plotArea[chartType], ["c:varyColors", "attrs", "val"]) : undefined;
    let dataPointStyles = [];
    if (chartType && plotArea[chartType]["c:ser"]) {
        const serArray = Array.isArray(plotArea[chartType]["c:ser"])
            ? plotArea[chartType]["c:ser"]
            : [plotArea[chartType]["c:ser"]];
        serArray.forEach(ser => {
            const dPtArray = ser["c:dPt"];
            if (dPtArray) {
                const dpStyles = {};
                const dpList = Array.isArray(dPtArray) ? dPtArray : [dPtArray];
                dpList.forEach(dp => {
                    const idx = dp["c:idx"]?.["attrs"]?.val;
                    const explosion = dp["c:explosion"]?.["attrs"]?.val;
                    const spPr = dp["c:spPr"];
                    if (idx !== undefined) {
                        const dpStyle = {};
                        if (explosion !== undefined) {
                            dpStyle.explosion = parseFloat(explosion);
                        }
                        if (spPr) {
                            const gradFill = spPr["a:gradFill"];
                            if (gradFill) {
                                dpStyle.gradientFill = PPTXStyleUtils.getGradientFill(gradFill, warpObj);
                            }
                        }
                        dpStyles[idx] = dpStyle;
                    }
                });
                dataPointStyles.push(dpStyles);
            }
        });
    }
    const chartTitleObj = PPTXStyleUtils.extractChartTitleStyle(chart, warpObj);
    const chartTitle = chartTitleObj.text;
    const chartStyle = {
        chartArea: PPTXStyleUtils.extractChartAreaStyle(chartSpace, warpObj),
        legend: PPTXStyleUtils.extractChartLegendStyle(chart, warpObj),
        categoryAxis: PPTXStyleUtils.extractChartAxisStyle(plotArea, "c:catAx", warpObj),
        valueAxis: PPTXStyleUtils.extractChartAxisStyle(plotArea, "c:valAx", warpObj),
        view3D: view3DProps,
        varyColors: varyColors === "1",
        dataPointStyles: dataPointStyles,
        title: chartTitleObj.style
    };
    let chartData = null;
    for (const key in plotArea) {
        switch (key) {
            case "c:lineChart":
                chartData = {
                    "type": "createChart",
                    "data": {
                        "chartId": `chart${warpObj.chartId.value++}`,
                        "chartType": "lineChart",
                        "chartData": PPTXStyleUtils.extractChartData(plotArea[key]["c:ser"], warpObj),
                        "style": chartStyle,
                        "title": chartTitle
                    }
                };
                warpObj.msgQueue.push(chartData);
                break;
            case "c:barChart":
                chartData = {
                    "type": "createChart",
                    "data": {
                        "chartId": `chart${warpObj.chartId.value++}`,
                        "chartType": "barChart",
                        "chartData": PPTXStyleUtils.extractChartData(plotArea[key]["c:ser"], warpObj),
                        "style": chartStyle,
                        "title": chartTitle
                    }
                };
                warpObj.msgQueue.push(chartData);
                break;
            case "c:pieChart":
                chartData = {
                    "type": "createChart",
                    "data": {
                        "chartId": `chart${warpObj.chartId.value++}`,
                        "chartType": "pieChart",
                        "chartData": PPTXStyleUtils.extractChartData(plotArea[key]["c:ser"], warpObj),
                        "style": chartStyle,
                        "title": chartTitle
                    }
                };
                warpObj.msgQueue.push(chartData);
                break;
            case "c:pie3DChart":
                chartData = {
                    "type": "createChart",
                    "data": {
                        "chartId": `chart${warpObj.chartId.value++}`,
                        "chartType": "pie3DChart",
                        "chartData": PPTXStyleUtils.extractChartData(plotArea[key]["c:ser"], warpObj),
                        "style": chartStyle,
                        "title": chartTitle
                    }
                };
                warpObj.msgQueue.push(chartData);
                break;
            case "c:areaChart":
                chartData = {
                    "type": "createChart",
                    "data": {
                        "chartId": `chart${warpObj.chartId.value++}`,
                        "chartType": "areaChart",
                        "chartData": PPTXStyleUtils.extractChartData(plotArea[key]["c:ser"], warpObj),
                        "style": chartStyle,
                        "title": chartTitle
                    }
                };
                warpObj.msgQueue.push(chartData);
                break;
            case "c:scatterChart":
                chartData = {
                    "type": "createChart",
                    "data": {
                        "chartId": `chart${warpObj.chartId.value++}`,
                        "chartType": "scatterChart",
                        "chartData": PPTXStyleUtils.extractChartData(plotArea[key]["c:ser"], warpObj),
                        "style": chartStyle,
                        "title": chartTitle
                    }
                };
                warpObj.msgQueue.push(chartData);
                break;
        }
    }
    return result;
}
function processMsgQueue(queue, result) {
    for (const msg of queue) {
        if (msg.type === "chart" || msg.type === "createChart") {
            const chartObj = msg.data;
            result.charts.push({
                chartId: chartObj.chartId,
                type: chartObj.chartType,
                data: chartObj.chartData,
                style: chartObj.style,
                title: chartObj.title
            });
        }
    }
}

async function genDiagram(node, wrapObj, source, shapeType, settings, parentNode) {
    node.attrs.order;
    const zip = wrapObj.zip;
    let xfrmNode = PPTXXmlUtils.getTextByPathList(node, ['p:xfrm']);
    const dgmRelIds = PPTXXmlUtils.getTextByPathList(node, ['a:graphic', 'a:graphicData', 'dgm:relIds', 'attrs']);
    const dgmClrFileId = dgmRelIds['r:cs'];
    const dgmDataFileId = dgmRelIds['r:dm'];
    const dgmLayoutFileId = dgmRelIds['r:lo'];
    const dgmQuickStyleFileId = dgmRelIds['r:qs'];
    const dgmClrFileName = wrapObj.slideResObj[dgmClrFileId].target;
    const dgmDataFileName = wrapObj.slideResObj[dgmDataFileId].target;
    const dgmLayoutFileName = wrapObj.slideResObj[dgmLayoutFileId].target;
    const dgmQuickStyleFileName = wrapObj.slideResObj[dgmQuickStyleFileId].target;
    await PPTXXmlUtils.readXmlFile(zip, dgmClrFileName);
    await PPTXXmlUtils.readXmlFile(zip, dgmDataFileName);
    await PPTXXmlUtils.readXmlFile(zip, dgmLayoutFileName);
    await PPTXXmlUtils.readXmlFile(zip, dgmQuickStyleFileName);
    const dgmDrwSpArray = PPTXXmlUtils.getTextByPathList(wrapObj.diagramContent, ['p:drawing', 'p:spTree', 'p:sp']);
    let result = '';
    if (dgmDrwSpArray !== undefined) {
        const results = [];
        for (const dspSp of dgmDrwSpArray) {
            PPTXXmlUtils.getTextByPathList(dspSp, ['p:txBody', 'a:p', 'a:r', 'a:t']);
            results.push(processSpNode(dspSp, node, wrapObj, 'diagramBg', shapeType));
        }
        const resolvedResults = await Promise.all(results);
        result = resolvedResults.join('');
    }
    let workingXfrmNode = xfrmNode;
    if (shapeType === 'group-abs' && wrapObj.currentGroupScale && xfrmNode) {
        const { scaleX, scaleY, childX, childY } = wrapObj.currentGroupScale;
        workingXfrmNode = JSON.parse(JSON.stringify(xfrmNode));
        if (xfrmNode['a:ext'] && xfrmNode['a:ext'].attrs) {
            const originalCx = parseInt(xfrmNode['a:ext'].attrs.cx);
            const originalCy = parseInt(xfrmNode['a:ext'].attrs.cy);
            workingXfrmNode['a:ext'].attrs.cx = Math.round(originalCx * scaleX);
            workingXfrmNode['a:ext'].attrs.cy = Math.round(originalCy * scaleY);
        }
        if (xfrmNode['a:off'] && xfrmNode['a:off'].attrs) {
            const originalOffX = parseInt(xfrmNode['a:off'].attrs.x);
            const originalOffY = parseInt(xfrmNode['a:off'].attrs.y);
            const relativeX = originalOffX - (childX / SLIDE_FACTOR$1);
            const relativeY = originalOffY - (childY / SLIDE_FACTOR$1);
            workingXfrmNode['a:off'].attrs.x = Math.round(childX / SLIDE_FACTOR$1 + relativeX * scaleX);
            workingXfrmNode['a:off'].attrs.y = Math.round(childY / SLIDE_FACTOR$1 + relativeY * scaleY);
        }
    }
    const position = PPTXXmlUtils.getPosition(workingXfrmNode, parentNode, undefined, undefined, shapeType);
    const size = PPTXXmlUtils.getSize(workingXfrmNode, undefined, undefined);
    let offX = 0, offY = 0, extCx = 0, extCy = 0;
    if (workingXfrmNode !== undefined) {
        if (workingXfrmNode['a:off'] && workingXfrmNode['a:off'].attrs) {
            offX = workingXfrmNode['a:off'].attrs.x || 0;
            offY = workingXfrmNode['a:off'].attrs.y || 0;
        }
        if (workingXfrmNode['a:ext'] && workingXfrmNode['a:ext'].attrs) {
            extCx = workingXfrmNode['a:ext'].attrs.cx || 0;
            extCy = workingXfrmNode['a:ext'].attrs.cy || 0;
        }
    }
    const dataAttrs = ` data-node-type="diagram" data-off-x="${offX}" data-off-y="${offY}" data-ext-cx="${extCx}" data-ext-cy="${extCy}"`;
    return `<div class='block diagram-content' style='${position}${size}'${dataAttrs}>${result}</div>`;
}
function indexNodes(content) {
    const keys = Object.keys(content);
    const spTreeNode = content[keys[0]]['p:cSld']['p:spTree'];
    const idTable = {};
    const idxTable = {};
    const typeTable = {};
    for (const key in spTreeNode) {
        if (key === 'p:nvGrpSpPr' || key === 'p:grpSpPr') {
            continue;
        }
        const targetNode = spTreeNode[key];
        if (Array.isArray(targetNode)) {
            for (const node of targetNode) {
                const nvSpPrNode = node['p:nvSpPr'];
                const id = PPTXXmlUtils.getTextByPathList(nvSpPrNode, ['p:cNvPr', 'attrs', 'id']);
                const idx = PPTXXmlUtils.getTextByPathList(nvSpPrNode, ['p:nvPr', 'p:ph', 'attrs', 'idx']);
                const type = PPTXXmlUtils.getTextByPathList(nvSpPrNode, ['p:nvPr', 'p:ph', 'attrs', 'type']);
                if (id !== undefined)
                    idTable[id] = node;
                if (idx !== undefined)
                    idxTable[idx] = node;
                if (type !== undefined)
                    typeTable[type] = node;
            }
        }
        else {
            const nvSpPrNode = targetNode['p:nvSpPr'];
            const id = PPTXXmlUtils.getTextByPathList(nvSpPrNode, ['p:cNvPr', 'attrs', 'id']);
            const idx = PPTXXmlUtils.getTextByPathList(nvSpPrNode, ['p:nvPr', 'p:ph', 'attrs', 'idx']);
            const type = PPTXXmlUtils.getTextByPathList(nvSpPrNode, ['p:nvPr', 'p:ph', 'attrs', 'type']);
            if (id !== undefined)
                idTable[id] = targetNode;
            if (idx !== undefined)
                idxTable[idx] = targetNode;
            if (type !== undefined)
                typeTable[type] = targetNode;
        }
    }
    return { idTable, idxTable, typeTable };
}
function objectToDataAttributes(obj, prefix = '') {
    if (!obj || typeof obj !== 'object') {
        return '';
    }
    let result = '';
    for (const key in obj) {
        if (obj.hasOwnProperty(key)) {
            const value = obj[key];
            const dataKey = prefix ? `${prefix}-${key}` : key;
            if (typeof value === 'object' && value !== null && !Array.isArray(value)) {
                result += objectToDataAttributes(value, dataKey);
            }
            else if (typeof value === 'string' || typeof value === 'number') {
                let attrValue;
                if (typeof value === 'number') {
                    attrValue = Math.round(value * 100) / 100;
                }
                else {
                    attrValue = value;
                }
                const escapedValue = String(attrValue).replace(/'/g, '&#39;').replace(/"/g, '&quot;');
                result += ` data-${dataKey}="${escapedValue}"`;
            }
        }
    }
    return result;
}
async function processGroupSpNode(node, parentNode, wrapObj, source, settings) {
    const xfrmNode = PPTXXmlUtils.getTextByPathList(node, ['p:grpSpPr', 'a:xfrm']);
    let groupStyle = '';
    let shapeType = 'group';
    let top, left, width, height;
    let rotate = 0;
    let x = 0, y = 0, cx = 0, cy = 0;
    let childX = 0, childY = 0, childCx = 0, childCy = 0;
    if (xfrmNode !== undefined) {
        x = Math.round(parseInt(xfrmNode['a:off'].attrs.x) * SLIDE_FACTOR$1 * 100) / 100;
        y = Math.round(parseInt(xfrmNode['a:off'].attrs.y) * SLIDE_FACTOR$1 * 100) / 100;
        let parentChOffX = 0, parentChOffY = 0;
        if (parentNode !== undefined) {
            const parentGrpXfrmNode = PPTXXmlUtils.getTextByPathList(parentNode, ['p:grpSpPr', 'a:xfrm']);
            if (parentGrpXfrmNode !== undefined && parentGrpXfrmNode['a:chOff'] !== undefined && parentGrpXfrmNode['a:chOff'].attrs !== undefined) {
                parentChOffX = Math.round(parseInt(parentGrpXfrmNode['a:chOff'].attrs.x) * SLIDE_FACTOR$1 * 100) / 100;
                parentChOffY = Math.round(parseInt(parentGrpXfrmNode['a:chOff'].attrs.y) * SLIDE_FACTOR$1 * 100) / 100;
            }
        }
        if (xfrmNode['a:chOff'] !== undefined && xfrmNode['a:chOff'].attrs !== undefined) {
            childX = Math.round(parseInt(xfrmNode['a:chOff'].attrs.x) * SLIDE_FACTOR$1 * 100) / 100;
            childY = Math.round(parseInt(xfrmNode['a:chOff'].attrs.y) * SLIDE_FACTOR$1 * 100) / 100;
        }
        else {
            childX = x;
            childY = y;
        }
        if (parentChOffX > 0 || parentChOffY > 0) {
            childX = childX - parentChOffX;
            childY = childY - parentChOffY;
            x = x - parentChOffX;
            y = y - parentChOffY;
        }
        cx = Math.round(parseInt(xfrmNode['a:ext'].attrs.cx) * SLIDE_FACTOR$1 * 100) / 100;
        cy = Math.round(parseInt(xfrmNode['a:ext'].attrs.cy) * SLIDE_FACTOR$1 * 100) / 100;
        if (xfrmNode['a:chExt'] !== undefined && xfrmNode['a:chExt'].attrs !== undefined) {
            childCx = Math.round(parseInt(xfrmNode['a:chExt'].attrs.cx) * SLIDE_FACTOR$1 * 100) / 100;
            childCy = Math.round(parseInt(xfrmNode['a:chExt'].attrs.cy) * SLIDE_FACTOR$1 * 100) / 100;
        }
        else {
            childCx = cx;
            childCy = cy;
        }
        rotate = parseInt(xfrmNode.attrs.rot) || 0;
        let rotationStyle = '';
        if (childCx > cx || childCy > cy) {
            const scaleX = childCx > 0 ? cx / childCx : 1;
            const scaleY = childCy > 0 ? cy / childCy : 1;
            wrapObj.currentGroupScale = { scaleX, scaleY, childX, childY };
            top = y;
            left = x;
            width = cx;
            height = cy;
            shapeType = 'group-abs';
        }
        else {
            wrapObj.currentGroupScale = null;
            top = y;
            left = x;
            width = cx;
            height = cy;
        }
        if (!isNaN(rotate)) {
            const degrees = PPTXXmlUtils.angleToDegrees(rotate);
            rotationStyle = `transform: rotate(${degrees}deg); transform-origin: center;`;
            if (degrees !== 0) {
                shapeType = 'group-rotate';
            }
        }
        if (rotationStyle)
            groupStyle += rotationStyle;
    }
    if (top !== undefined)
        groupStyle += `top: ${top}px;`;
    if (left !== undefined)
        groupStyle += `left: ${left}px;`;
    if (width !== undefined)
        groupStyle += `width: ${width}px;`;
    if (height !== undefined)
        groupStyle += `height: ${height}px;`;
    const order = node.attrs.order;
    const dataAttrs = objectToDataAttributes({
        'node-id': PPTXXmlUtils.getTextByPathList(node, ['p:nvGrpSpPr', 'p:cNvPr', 'attrs', 'id']),
        'node-name': PPTXXmlUtils.getTextByPathList(node, ['p:nvGrpSpPr', 'p:cNvPr', 'attrs', 'name']),
        'off-x': x,
        'off-y': y,
        'ext-cx': cx,
        'ext-cy': cy,
        'ch-off-x': childX,
        'ch-off-y': childY,
        'ch-ext-cx': childCx,
        'ch-ext-cy': childCy,
        'shape-type': shapeType,
        'rotate': rotate
    });
    let result = `<div class='block group' style='z-index: ${order};${groupStyle}'${dataAttrs}>`;
    const previousGroupScale = wrapObj.currentGroupScale;
    for (const nodeKey in node) {
        if (Array.isArray(node[nodeKey])) {
            for (const childNode of node[nodeKey]) {
                result += await processNodesInSlide(nodeKey, childNode, node, wrapObj, source, shapeType, settings, node);
            }
        }
        else {
            result += await processNodesInSlide(nodeKey, node[nodeKey], node, wrapObj, source, shapeType, settings, node);
        }
    }
    wrapObj.currentGroupScale = previousGroupScale;
    result += '</div>';
    return result;
}
async function processNodesInSlide(nodeKey, nodeValue, nodes, wrapObj, source, shapeType, settings, parentNode) {
    switch (nodeKey) {
        case 'p:sp':
            return await processSpNode(nodeValue, parentNode, wrapObj, source, shapeType, settings);
        case 'p:cxnSp':
            return await processCxnSpNode(nodeValue, parentNode, wrapObj, source, shapeType, settings);
        case 'p:pic':
            return await processPicNode(nodeValue, parentNode, wrapObj, source, shapeType, settings);
        case 'p:graphicFrame':
            return await processGraphicFrameNode(nodeValue, parentNode, wrapObj, source, shapeType, settings);
        case 'p:grpSp':
            return await processGroupSpNode(nodeValue, parentNode, wrapObj, source, settings);
        case 'mc:AlternateContent':
            const mcFallbackNode = PPTXXmlUtils.getTextByPathList(nodeValue, ['mc:Fallback']);
            return await processGroupSpNode(mcFallbackNode, parentNode, wrapObj, source, settings);
        default:
            return '';
    }
}
async function processSpNode(node, parentNode, wrapObj, source, shapeType, settings) {
    const id = PPTXXmlUtils.getTextByPathList(node, ['p:nvSpPr', 'p:cNvPr', 'attrs', 'id']);
    const name = PPTXXmlUtils.getTextByPathList(node, ['p:nvSpPr', 'p:cNvPr', 'attrs', 'name']);
    let idx = PPTXXmlUtils.getTextByPathList(node, ['p:nvSpPr', 'p:nvPr', 'p:ph', 'attrs', 'idx']);
    let type = PPTXXmlUtils.getTextByPathList(node, ['p:nvSpPr', 'p:nvPr', 'p:ph', 'attrs', 'type']);
    const order = PPTXXmlUtils.getTextByPathList(node, ['attrs', 'order']);
    let isUserDrawnBg;
    if (source === 'slideLayoutBg' || source === 'slideMasterBg') {
        const userDrawn = PPTXXmlUtils.getTextByPathList(node, ['p:nvSpPr', 'p:nvPr', 'attrs', 'userDrawn']);
        isUserDrawnBg = userDrawn === '1';
    }
    let slideLayoutSpNode;
    let slideMasterSpNode;
    if (idx !== undefined) {
        slideLayoutSpNode = wrapObj.slideLayoutTables.idxTable[idx];
        if (type !== undefined) {
            slideMasterSpNode = wrapObj.slideMasterTables.typeTable[type];
        }
        else {
            slideMasterSpNode = wrapObj.slideMasterTables.idxTable[idx];
        }
    }
    else if (type !== undefined) {
        slideLayoutSpNode = wrapObj.slideLayoutTables.typeTable[type];
        slideMasterSpNode = wrapObj.slideMasterTables.typeTable[type];
    }
    if (type === undefined) {
        const txBoxVal = PPTXXmlUtils.getTextByPathList(node, ['p:nvSpPr', 'p:cNvSpPr', 'attrs', 'txBox']);
        if (txBoxVal === '1') {
            type = 'textBox';
        }
    }
    if (type === undefined) {
        type = PPTXXmlUtils.getTextByPathList(slideLayoutSpNode, ['p:nvSpPr', 'p:nvPr', 'p:ph', 'attrs', 'type']);
        if (type === undefined) {
            type = source === 'diagramBg' ? 'diagram' : 'obj';
        }
    }
    const result = await PPTXShapeUtils.genShape(node, parentNode, slideLayoutSpNode, slideMasterSpNode, id, name, idx, type, order, wrapObj, isUserDrawnBg, shapeType, source, settings);
    return result;
}
async function processCxnSpNode(node, parentNode, wrapObj, source, shapeType, settings) {
    const id = node['p:nvCxnSpPr']['p:cNvPr'].attrs.id;
    const name = node['p:nvCxnSpPr']['p:cNvPr'].attrs.name;
    const idx = node['p:nvCxnSpPr']['p:nvPr']['p:ph'] === undefined
        ? undefined
        : node['p:nvCxnSpPr']['p:nvPr']['p:ph'].attrs.idx;
    const type = node['p:nvCxnSpPr']['p:nvPr']['p:ph'] === undefined
        ? undefined
        : node['p:nvCxnSpPr']['p:nvPr']['p:ph'].attrs.type;
    const order = node.attrs.order;
    return await PPTXShapeUtils.genShape(node, parentNode, undefined, undefined, id, name, idx, type, order, wrapObj, undefined, shapeType, source, settings);
}
async function processPicNode(node, parentNode, wrapObj, source, shapeType, settings) {
    const order = node.attrs.order;
    const rid = node['p:blipFill']['a:blip'].attrs['r:embed'];
    let resObj;
    if (source === 'slideMasterBg') {
        resObj = wrapObj.masterResObj;
    }
    else if (source === 'slideLayoutBg') {
        resObj = wrapObj.layoutResObj;
    }
    else {
        resObj = wrapObj.slideResObj;
    }
    if (resObj === undefined) {
        resObj = wrapObj.slideResObj;
    }
    if (resObj === undefined) {
        return '';
    }
    const imgName = resObj[rid]?.target;
    if (imgName === undefined) {
        return '';
    }
    const imgFileExt = PPTXXmlUtils.extractFileExtension(imgName).toLowerCase();
    const zip = wrapObj.zip;
    let context = 'slide';
    if (source === 'slideMasterBg') {
        context = 'master';
    }
    else if (source === 'slideLayoutBg') {
        context = 'layout';
    }
    const imgFile = PPTXXmlUtils.findMediaFile(zip, imgName, context, '');
    if (imgFile === null) {
        return '';
    }
    const imgArrayBuffer = await imgFile.async("arraybuffer");
    let xfrmNode = node['p:spPr']?.['a:xfrm'];
    if (xfrmNode === undefined) {
        const idx = PPTXXmlUtils.getTextByPathList(node, ['p:nvPicPr', 'p:nvPr', 'p:ph', 'attrs', 'idx']);
        if (idx !== undefined) {
            xfrmNode = PPTXXmlUtils.getTextByPathList(wrapObj.slideLayoutTables, ['idxTable', idx, 'p:spPr', 'a:xfrm']);
        }
    }
    let rotate = 0;
    const rotateNode = PPTXXmlUtils.getTextByPathList(node, ['p:spPr', 'a:xfrm', 'attrs', 'rot']);
    if (rotateNode !== undefined) {
        rotate = PPTXXmlUtils.angleToDegrees(rotateNode);
    }
    let mediaSupportFlag = false;
    let mediaPicFlag = false;
    let isVideoLink = false;
    let videoBlob, videoFile;
    const vdoNode = PPTXXmlUtils.getTextByPathList(node, ['p:nvPicPr', 'p:nvPr', 'a:videoFile']);
    const mediaProcess = settings.mediaProcess;
    if (vdoNode !== undefined && mediaProcess) {
        const vdoRid = vdoNode.attrs['r:link'];
        videoFile = resObj[vdoRid].target;
        const checkIfLink = PPTXXmlUtils.IsVideoLink(videoFile);
        if (checkIfLink) {
            videoFile = PPTXXmlUtils.convertVideoToEmbed(videoFile);
            videoFile = PPTXXmlUtils.escapeHtml(videoFile);
            isVideoLink = true;
            mediaSupportFlag = true;
            mediaPicFlag = true;
        }
        else {
            const vdoFileExt = PPTXXmlUtils.extractFileExtension(videoFile).toLowerCase();
            if (['mp4', 'webm', 'ogg'].includes(vdoFileExt)) {
                const vdoFileObj = PPTXXmlUtils.findMediaFile(zip, videoFile, context, '');
                if (vdoFileObj !== null) {
                    const uInt8Array = await vdoFileObj.async("arraybuffer");
                    const vdoMimeType = PPTXXmlUtils.getMimeType(vdoFileExt);
                    const blob = new Blob([uInt8Array], { type: vdoMimeType });
                    videoBlob = URL.createObjectURL(blob);
                    mediaSupportFlag = true;
                    mediaPicFlag = true;
                }
            }
        }
    }
    let audioPlayerFlag = false;
    let audioBlob;
    let audioObj;
    const audioNode = PPTXXmlUtils.getTextByPathList(node, ['p:nvPicPr', 'p:nvPr', 'a:audioFile']);
    if (audioNode !== undefined && mediaProcess) {
        const audioRid = audioNode.attrs['r:link'];
        const audioFile = resObj[audioRid].target;
        const audioFileExt = PPTXXmlUtils.extractFileExtension(audioFile).toLowerCase();
        if (['mp3', 'wav', 'ogg'].includes(audioFileExt)) {
            const audioFileObj = PPTXXmlUtils.findMediaFile(zip, audioFile, context, '');
            if (audioFileObj !== null) {
                const uInt8ArrayAudio = await audioFileObj.async("arraybuffer");
                const blobAudio = new Blob([uInt8ArrayAudio]);
                audioBlob = URL.createObjectURL(blobAudio);
                const cx = parseInt(xfrmNode['a:ext'].attrs.cx) * 20;
                const cy = parseInt(xfrmNode['a:ext'].attrs.cy);
                const x = parseInt(xfrmNode['a:off'].attrs.x) / 2.5;
                const y = parseInt(xfrmNode['a:off'].attrs.y);
                audioObj = {
                    'a:ext': { attrs: { cx, cy } },
                    'a:off': { attrs: { x, y } }
                };
                audioPlayerFlag = true;
                mediaSupportFlag = true;
                mediaPicFlag = true;
            }
        }
    }
    const mimeType = PPTXXmlUtils.getMimeType(imgFileExt);
    let scaledXfrmNode = null;
    if (shapeType === 'group-abs' && wrapObj.currentGroupScale) {
        const { scaleX, scaleY, childX, childY } = wrapObj.currentGroupScale;
        if (xfrmNode !== undefined) {
            scaledXfrmNode = JSON.parse(JSON.stringify(xfrmNode));
            if (xfrmNode['a:ext'] && xfrmNode['a:ext'].attrs) {
                const originalCx = parseInt(xfrmNode['a:ext'].attrs.cx);
                const originalCy = parseInt(xfrmNode['a:ext'].attrs.cy);
                scaledXfrmNode['a:ext'].attrs.cx = Math.round(originalCx * scaleX);
                scaledXfrmNode['a:ext'].attrs.cy = Math.round(originalCy * scaleY);
            }
            if (xfrmNode['a:off'] && xfrmNode['a:off'].attrs) {
                const originalOffX = parseInt(xfrmNode['a:off'].attrs.x);
                const originalOffY = parseInt(xfrmNode['a:off'].attrs.y);
                const relativeX = originalOffX - (childX / SLIDE_FACTOR$1);
                const relativeY = originalOffY - (childY / SLIDE_FACTOR$1);
                scaledXfrmNode['a:off'].attrs.x = Math.round(childX / SLIDE_FACTOR$1 + relativeX * scaleX);
                scaledXfrmNode['a:off'].attrs.y = Math.round(childY / SLIDE_FACTOR$1 + relativeY * scaleY);
            }
        }
    }
    const position = mediaProcess && audioPlayerFlag
        ? PPTXXmlUtils.getPosition(audioObj, parentNode, undefined, undefined, shapeType)
        : PPTXXmlUtils.getPosition(scaledXfrmNode || xfrmNode, parentNode, undefined, undefined, shapeType);
    const size = mediaProcess && audioPlayerFlag
        ? PPTXXmlUtils.getSize(audioObj, undefined, undefined)
        : PPTXXmlUtils.getSize(scaledXfrmNode || xfrmNode, undefined, undefined);
    let imgOffX = 0, imgOffY = 0, imgExtCx = 0, imgExtCy = 0;
    if (xfrmNode !== undefined) {
        if (xfrmNode['a:off'] && xfrmNode['a:off'].attrs) {
            imgOffX = parseInt(xfrmNode['a:off'].attrs.x) * SLIDE_FACTOR$1;
            imgOffY = parseInt(xfrmNode['a:off'].attrs.y) * SLIDE_FACTOR$1;
        }
        if (xfrmNode['a:ext'] && xfrmNode['a:ext'].attrs) {
            imgExtCx = parseInt(xfrmNode['a:ext'].attrs.cx) * SLIDE_FACTOR$1;
            imgExtCy = parseInt(xfrmNode['a:ext'].attrs.cy) * SLIDE_FACTOR$1;
        }
    }
    const dataAttrs = objectToDataAttributes({
        'node-id': PPTXXmlUtils.getTextByPathList(node, ['p:nvPicPr', 'p:cNvPr', 'attrs', 'id']),
        'node-name': PPTXXmlUtils.getTextByPathList(node, ['p:nvPicPr', 'p:cNvPr', 'attrs', 'name']),
        'node-descr': PPTXXmlUtils.getTextByPathList(node, ['p:nvPicPr', 'p:cNvPr', 'attrs', 'descr']),
        'off-x': imgOffX,
        'off-y': imgOffY,
        'ext-cx': imgExtCx,
        'ext-cy': imgExtCy,
        'shape-type': shapeType,
        'rotate': rotate,
        'is-video': (vdoNode !== undefined) ? 'true' : 'false',
        'is-audio': (audioNode !== undefined) ? 'true' : 'false',
        'is-gif': (mimeType === 'image/gif') ? 'true' : 'false'
    });
    let result = `<div class='block content' style='${position}${size} z-index: ${order};transform: rotate(${rotate}deg);'${dataAttrs}>`;
    if ((vdoNode === undefined && audioNode === undefined) || !mediaProcess || !mediaSupportFlag) {
        const base64Data = PPTXXmlUtils.base64ArrayBuffer(imgArrayBuffer);
        const gifAttrs = (mimeType === 'image/gif') ? 'autoplay loop muted playsinline' : '';
        result += `<img src='data:${mimeType};base64,${base64Data}' style='width: 100%; height: 100%' ${gifAttrs}/>`;
    }
    else if ((vdoNode !== undefined || audioNode !== undefined) && mediaProcess && mediaSupportFlag) {
        if (vdoNode !== undefined && !isVideoLink) {
            result += `<video src='${videoBlob}' autoplay loop muted controls style='width: 100%; height: 100%'>Your browser does not support the video tag.</video>`;
        }
        else if (vdoNode !== undefined && isVideoLink) {
            const iframeAttrs = 'allowfullscreen allow="accelerometer; autoplay; clipboard-write; encrypted-media; gyroscope; picture-in-picture" loading="lazy"';
            result += `<iframe src='${videoFile}' ${iframeAttrs} style='width: 100%; height: 100%; border: none;'></iframe>`;
        }
        if (audioNode !== undefined) {
            result += `<audio id="audio_player" controls><source src="${audioBlob}"></audio>`;
        }
    }
    if (!mediaSupportFlag && mediaPicFlag) {
        result += `<span style='color:red;font-size:40px;position: absolute;'>This media file Not supported by HTML5</span>`;
    }
    result += '</div>';
    return result;
}
async function processGraphicFrameNode(node, parentNode, wrapObj, source, shapeType, settings) {
    const graphicTypeUri = PPTXXmlUtils.getTextByPathList(node, ['a:graphic', 'a:graphicData', 'attrs', 'uri']);
    switch (graphicTypeUri) {
        case 'http://schemas.openxmlformats.org/drawingml/2006/table':
            return await PPTXTextUtils.genTable(node, wrapObj, shapeType);
        case 'http://schemas.openxmlformats.org/drawingml/2006/chart':
            return await genChart(node, wrapObj, parentNode);
        case 'http://schemas.openxmlformats.org/drawingml/2006/diagram':
            return await genDiagram(node, wrapObj, source, shapeType, settings, parentNode);
        case 'http://schemas.openxmlformats.org/presentationml/2006/ole':
            let oleObjNode = PPTXXmlUtils.getTextByPathList(node, ['a:graphic', 'a:graphicData', 'mc:AlternateContent', 'mc:Fallback', 'p:oleObj']);
            if (oleObjNode === undefined) {
                oleObjNode = PPTXXmlUtils.getTextByPathList(node, ['a:graphic', 'a:graphicData', 'p:oleObj']);
            }
            if (oleObjNode !== undefined) {
                return await processGroupSpNode(oleObjNode, undefined, wrapObj, source, settings);
            }
            return '';
        default:
            return '';
    }
}
function processSpPrNode(node, wrapObj) {
}
async function getBackground(wrapObj, slideSize, index, settings) {
    const { slideContent, slideLayoutContent, slideMasterContent } = wrapObj;
    const nodesSldLayout = PPTXXmlUtils.getTextByPathList(slideLayoutContent, ['p:sldLayout', 'p:cSld', 'p:spTree']);
    const nodesSldMaster = PPTXXmlUtils.getTextByPathList(slideMasterContent, ['p:sldMaster', 'p:cSld', 'p:spTree']);
    const showMasterSp = PPTXXmlUtils.getTextByPathList(slideLayoutContent, ['p:sldLayout', 'attrs', 'showMasterSp']);
    const bgColor = await PPTXStyleUtils.getSlideBackgroundFill(wrapObj, index);
    let result = `<div class='slide-background-${index}' style='width:${slideSize.width}px; height:${slideSize.height}px;${bgColor}'>`;
    if (nodesSldLayout !== undefined) {
        for (const nodeKey in nodesSldLayout) {
            if (Array.isArray(nodesSldLayout[nodeKey])) {
                for (const node of nodesSldLayout[nodeKey]) {
                    const phType = PPTXXmlUtils.getTextByPathList(node, ['p:nvSpPr', 'p:nvPr', 'p:ph', 'attrs', 'type']);
                    if (phType !== 'pic') {
                        result += await processNodesInSlide(nodeKey, node, nodesSldLayout, wrapObj, 'slideLayoutBg', 'group', settings, undefined);
                    }
                }
            }
            else {
                const phType = PPTXXmlUtils.getTextByPathList(nodesSldLayout[nodeKey], ['p:nvSpPr', 'p:nvPr', 'p:ph', 'attrs', 'type']);
                if (phType !== 'pic') {
                    result += await processNodesInSlide(nodeKey, nodesSldLayout[nodeKey], nodesSldLayout, wrapObj, 'slideLayoutBg', 'group', settings, undefined);
                }
            }
        }
    }
    if (nodesSldMaster !== undefined && (showMasterSp === '1' || showMasterSp === undefined)) {
        for (const nodeKey in nodesSldMaster) {
            if (Array.isArray(nodesSldMaster[nodeKey])) {
                for (const node of nodesSldMaster[nodeKey]) {
                    PPTXXmlUtils.getTextByPathList(node, ['p:nvSpPr', 'p:nvPr', 'p:ph', 'attrs', 'type']);
                    result += await processNodesInSlide(nodeKey, node, nodesSldMaster, wrapObj, 'slideMasterBg', 'group', settings, undefined);
                }
            }
            else {
                result += await processNodesInSlide(nodeKey, nodesSldMaster[nodeKey], nodesSldMaster, wrapObj, 'slideMasterBg', 'group', settings, undefined);
            }
        }
    }
    return result;
}
const PPTXNodeUtils = {
    indexNodes,
    processGroupSpNode,
    processNodesInSlide,
    processSpNode,
    processCxnSpNode,
    processPicNode,
    processGraphicFrameNode,
    processSpPrNode,
    getBackground,
    genDiagram
};

const PX_TO_EMU = 914400 / 96;
const PT_TO_EMU = 12700;
const NS = {
    a: 'http://schemas.openxmlformats.org/drawingml/2006/main',
    r: 'http://schemas.openxmlformats.org/officeDocument/2006/relationships',
    p: 'http://schemas.openxmlformats.org/presentationml/2006/main',
    c: 'http://schemas.openxmlformats.org/drawingml/2006/chart',
    rel: 'http://schemas.openxmlformats.org/package/2006/relationships',
    cp: 'http://schemas.openxmlformats.org/package/2006/metadata/core-properties',
    dc: 'http://purl.org/dc/elements/1.1/',
    dcterms: 'http://purl.org/dc/terms/',
    dcmitype: 'http://purl.org/dc/dcmitype/',
    xsi: 'http://www.w3.org/2001/XMLSchema-instance',
    ext: 'http://schemas.openxmlformats.org/officeDocument/2006/extended-properties'
};
function escapeXml(str) {
    if (str === undefined || str === null)
        return '';
    return String(str)
        .replace(/&/g, '&amp;')
        .replace(/</g, '&lt;')
        .replace(/>/g, '&gt;')
        .replace(/"/g, '&quot;')
        .replace(/'/g, '&apos;');
}
function pxToEmu(px) {
    return Math.round((Number(px) || 0) * PX_TO_EMU);
}
function ptToEmu(pt) {
    return Math.round((Number(pt) || 0) * PT_TO_EMU);
}
function ptToSz(pt) {
    return Math.round((Number(pt) || 0) * 100);
}
function degToRot(deg) {
    return Math.round((Number(deg) || 0) * 60000);
}
function colorToHex(color) {
    if (color === undefined || color === null || color === '')
        return '000000';
    const c = new tinycolor$2(String(color));
    if (!c.isValid())
        return '000000';
    return c.toHexString().replace('#', '').toUpperCase();
}
function xmlNode(tagName, attrs, ...children) {
    if (attrs && typeof attrs === 'object' && !Array.isArray(attrs) && typeof attrs.tagName === 'string') {
        children.unshift(attrs);
        attrs = null;
    }
    const filteredAttrs = {};
    if (attrs) {
        for (const key in attrs) {
            const val = attrs[key];
            if (val !== undefined && val !== null) {
                filteredAttrs[key] = String(val);
            }
        }
    }
    return {
        tagName,
        attrs: filteredAttrs,
        children: children.filter(c => c !== undefined && c !== null && c !== '')
    };
}
function nodeToString(node, indent = '') {
    if (typeof node === 'string') {
        return escapeXml(node);
    }
    const attrs = Object.keys(node.attrs)
        .map(key => ` ${key}="${escapeXml(node.attrs[key])}"`)
        .join('');
    if (!node.children || node.children.length === 0) {
        return `<${node.tagName}${attrs}/>`;
    }
    const inner = node.children
        .map((child) => nodeToString(child))
        .join('');
    if (node.children.every((c) => typeof c === 'string')) {
        return `<${node.tagName}${attrs}>${inner}</${node.tagName}>`;
    }
    return `<${node.tagName}${attrs}>${inner}</${node.tagName}>`;
}
function toXmlDocument(rootNode) {
    return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\r\n${nodeToString(rootNode)}`;
}

const DRAWING_NS = `xmlns:a="${NS.a}" xmlns:r="${NS.r}" xmlns:p="${NS.p}"`;
function buildThemeXml() {
    return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\r\n<a:theme xmlns:a="${NS.a}" name="Office Theme"><a:themeElements><a:clrScheme name="Office"><a:dk1><a:sysClr val="windowText" lastClr="000000"/></a:dk1><a:lt1><a:sysClr val="window" lastClr="FFFFFF"/></a:lt1><a:dk2><a:srgbClr val="44546A"/></a:dk2><a:lt2><a:srgbClr val="E7E6E6"/></a:lt2><a:accent1><a:srgbClr val="4472C4"/></a:accent1><a:accent2><a:srgbClr val="ED7D31"/></a:accent2><a:accent3><a:srgbClr val="A5A5A5"/></a:accent3><a:accent4><a:srgbClr val="FFC000"/></a:accent4><a:accent5><a:srgbClr val="5B9BD5"/></a:accent5><a:accent6><a:srgbClr val="70AD47"/></a:accent6><a:hlink><a:srgbClr val="0563C1"/></a:hlink><a:folHlink><a:srgbClr val="954F72"/></a:folHlink></a:clrScheme><a:fontScheme name="Office"><a:majorFont><a:latin typeface="Calibri Light"/><a:ea typeface=""/><a:cs typeface=""/></a:majorFont><a:minorFont><a:latin typeface="Calibri"/><a:ea typeface=""/><a:cs typeface=""/></a:minorFont></a:fontScheme><a:fmtScheme name="Office"><a:fillStyleLst><a:solidFill><a:schemeClr val="phClr"/></a:solidFill><a:gradFill rotWithShape="1"><a:gsLst><a:gs pos="0"><a:schemeClr val="phClr"><a:tint val="94000"/><a:satMod val="110000"/></a:schemeClr></a:gs><a:gs pos="1000"><a:schemeClr val="phClr"><a:tint val="94000"/><a:satMod val="120000"/></a:schemeClr></a:gs><a:gs pos="100000"><a:schemeClr val="phClr"><a:shade val="94000"/><a:satMod val="120000"/></a:schemeClr></a:gs></a:gsLst><a:lin ang="4553000" scaled="0"/></a:gradFill><a:solidFill><a:schemeClr val="phClr"><a:tint val="60000"/><a:satMod val="170000"/></a:schemeClr></a:solidFill></a:fillStyleLst><a:lnStyleLst><a:ln w="6350" cap="flat" cmpd="sng" algn="ctr"><a:solidFill><a:schemeClr val="phClr"/></a:solidFill><a:prstDash val="solid"/><a:miter lim="800000"/></a:ln><a:ln w="12700" cap="flat" cmpd="sng" algn="ctr"><a:solidFill><a:schemeClr val="phClr"/></a:solidFill><a:prstDash val="solid"/><a:miter lim="800000"/></a:ln><a:ln w="19050" cap="flat" cmpd="sng" algn="ctr"><a:solidFill><a:schemeClr val="phClr"/></a:solidFill><a:prstDash val="solid"/><a:miter lim="800000"/></a:ln></a:lnStyleLst><a:effectStyleLst><a:effectStyle><a:effectLst/></a:effectStyle><a:effectStyle><a:effectLst/></a:effectStyle><a:effectStyle><a:effectLst><a:outerShdw blurRad="57150" dist="19050" dir="5400000" algn="ctr" rotWithShape="0"><a:srgbClr val="000000"><a:alpha val="63000"/></a:srgbClr></a:outerShdw></a:effectLst></a:effectStyle></a:effectStyleLst><a:bgFillStyleLst><a:solidFill><a:schemeClr val="phClr"/></a:solidFill><a:solidFill><a:schemeClr val="phClr"><a:tint val="95000"/><a:satMod val="170000"/></a:schemeClr></a:solidFill><a:gradFill rotWithShape="1"><a:gsLst><a:gs pos="0"><a:schemeClr val="phClr"><a:tint val="93000"/><a:satMod val="150000"/></a:schemeClr></a:gs><a:gs pos="100000"><a:schemeClr val="phClr"><a:shade val="97000"/><a:satMod val="130000"/></a:schemeClr></a:gs></a:gsLst><a:lin ang="5400000" scaled="0"/></a:gradFill></a:bgFillStyleLst></a:fmtScheme></a:themeElements></a:theme>`;
}
function buildSlideMasterXml() {
    return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\r\n<p:sldMaster ${DRAWING_NS}><p:cSld><p:bg><p:bgRef idx="1001"><a:schemeClr val="bg1"/></p:bgRef></p:bg><p:spTree><p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr><p:grpSpPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="0" cy="0"/><a:chOff x="0" y="0"/><a:chExt cx="0" cy="0"/></a:xfrm></p:grpSpPr></p:spTree></p:cSld><p:clrMap bg1="lt1" tx1="dk1" bg2="lt2" tx2="dk2" accent1="accent1" accent2="accent2" accent3="accent3" accent4="accent4" accent5="accent5" accent6="accent6" hlink="hlink" folHlink="folHlink"/><p:sldLayoutIdLst><p:sldLayoutId id="2147483649" r:id="rId1"/></p:sldLayoutIdLst><p:txStyles><p:titleStyle><a:lvl1pPr><a:defRPr sz="4400"/></a:lvl1pPr></p:titleStyle><p:bodyStyle><a:lvl1pPr><a:defRPr sz="3200"/></a:lvl1pPr></p:bodyStyle><p:otherStyle><a:lvl1pPr><a:defRPr sz="1800"/></a:lvl1pPr></p:otherStyle></p:txStyles></p:sldMaster>`;
}
function buildSlideLayoutXml() {
    return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\r\n<p:sldLayout ${DRAWING_NS} type="blank" preserve="1"><p:cSld name="Blank"><p:spTree><p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr><p:grpSpPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="0" cy="0"/><a:chOff x="0" y="0"/><a:chExt cx="0" cy="0"/></a:xfrm></p:grpSpPr></p:spTree></p:cSld><p:clrMapOvr><a:masterClrMapping/></p:clrMapOvr></p:sldLayout>`;
}
function buildPresPropsXml() {
    return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\r\n<p:presentationPr xmlns:a="${NS.a}" xmlns:r="${NS.r}" xmlns:p="${NS.p}"/>`;
}
function buildViewPropsXml() {
    return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\r\n<p:viewProps xmlns:a="${NS.a}" xmlns:r="${NS.r}" xmlns:p="${NS.p}"/>`;
}
function buildTableStylesXml() {
    return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\r\n<p:tblStyLst xmlns:a="${NS.a}" xmlns:p="${NS.p}" def="{5C22544A-7EE6-4342-B048-85BDC9FD1C3A}"/>`;
}
function buildPresentationXml(slideSize, slides) {
    const slideEntries = slides
        .map((s, i) => `<p:sldId id="${256 + i}" r:id="${escapeXml(s.relId)}"/>`)
        .join('');
    return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\r\n<p:presentation xmlns:a="${NS.a}" xmlns:r="${NS.r}" xmlns:p="${NS.p}" saveSubsetFonts="1"><p:sldMasterIdLst><p:sldMasterId id="2147483648" r:id="rId1"/></p:sldMasterIdLst><p:sldIdLst>${slideEntries}</p:sldIdLst><p:sldSz cx="${pxToEmu(slideSize.width)}" cy="${pxToEmu(slideSize.height)}"/><p:notesSz cx="6858000" cy="914400"/><p:defaultTextStyle/></p:presentation>`;
}
function buildRelationshipsXml(rels) {
    const entries = rels
        .map((r) => `<Relationship Id="${escapeXml(r.relId)}" Type="${escapeXml(r.type)}" Target="${escapeXml(r.target)}"${r.external ? ' TargetMode="External"' : ''}/>`)
        .join('');
    return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\r\n<Relationships xmlns="${NS.rel}">${entries}</Relationships>`;
}
function buildContentTypesXml(mediaExts, slideCount) {
    const MIME_MAP = {
        png: 'image/png',
        jpeg: 'image/jpeg',
        jpg: 'image/jpeg',
        gif: 'image/gif',
        bmp: 'image/bmp',
        svg: 'image/svg+xml',
        tiff: 'image/tiff',
        webp: 'image/webp',
        emf: 'image/x-emf',
        wmf: 'image/x-wmf'
    };
    const defaults = ['rels', 'xml']
        .map(ext => `<Default Extension="${ext}" ContentType="${ext === 'rels' ? 'application/vnd.openxmlformats-package.relationships+xml' : 'application/xml'}"/>`)
        .join('');
    const mediaDefaults = [...new Set(mediaExts)]
        .map(ext => `<Default Extension="${escapeXml(ext)}" ContentType="${MIME_MAP[String(ext).toLowerCase()] || 'application/octet-stream'}"/>`)
        .join('');
    const slideOverrides = Array.from({ length: slideCount }, (_, i) => `<Override PartName="/ppt/slides/slide${i + 1}.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slide+xml"/>`).join('');
    return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\r\n<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">${defaults}${mediaDefaults}<Override PartName="/ppt/presentation.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.presentation.main+xml"/><Override PartName="/ppt/slideMasters/slideMaster1.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slideMaster+xml"/><Override PartName="/ppt/slideLayouts/slideLayout1.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slideLayout+xml"/><Override PartName="/ppt/theme/theme1.xml" ContentType="application/vnd.openxmlformats-officedocument.theme+xml"/><Override PartName="/docProps/core.xml" ContentType="application/vnd.openxmlformats-package.core-properties+xml"/><Override PartName="/docProps/app.xml" ContentType="application/vnd.openxmlformats-officedocument.extended-properties+xml"/>${slideOverrides}</Types>`;
}
function buildCorePropsXml(metadata) {
    const md = metadata || {};
    const now = new Date().toISOString().replace(/\.\d+Z$/, 'Z');
    const created = md.created || now;
    const modified = md.modified || now;
    const fields = [
        ['dc:title', md.title],
        ['dc:subject', md.subject],
        ['dc:creator', md.author],
        ['cp:keywords', md.keywords],
        ['dc:description', md.description],
        ['cp:lastModifiedBy', md.lastModifiedBy || md.author],
        ['cp:category', md.category],
        ['cp:contentStatus', md.status]
    ]
        .filter(([, val]) => val !== undefined && val !== null && val !== '')
        .map(([tag, val]) => `<${tag}>${escapeXml(val)}</${tag}>`)
        .join('');
    return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\r\n<cp:coreProperties xmlns:cp="${NS.cp}" xmlns:dc="${NS.dc}" xmlns:dcterms="${NS.dcterms}" xmlns:dcmitype="${NS.dcmitype}" xmlns:xsi="${NS.xsi}">${fields}<dcterms:created xsi:type="dcterms:W3CDTF">${created}</dcterms:created><dcterms:modified xsi:type="dcterms:W3CDTF">${modified}</dcterms:modified></cp:coreProperties>`;
}
function buildAppPropsXml(slideCount) {
    return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\r\n<Properties xmlns="${NS.ext}" xmlns:vt="http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes"><Application>Microsoft Office PowerPoint</Application><Slides>${slideCount}</Slides><AppVersion>16.0000</AppVersion></Properties>`;
}
function buildRootRelsXml() {
    return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\r\n<Relationships xmlns="${NS.rel}"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="ppt/presentation.xml"/><Relationship Id="rId2" Type="http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties" Target="docProps/core.xml"/><Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/extended-properties" Target="docProps/app.xml"/></Relationships>`;
}
const MASTER_RELS = [
    { relId: 'rId1', type: 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideLayout', target: '../slideLayouts/slideLayout1.xml' },
    { relId: 'rId2', type: 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme', target: '../theme/theme1.xml' }
];
const LAYOUT_RELS = [
    { relId: 'rId1', type: 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideMaster', target: '../slideMasters/slideMaster1.xml' }
];
const REL_TYPES = {
    slide: 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide',
    slideLayout: 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideLayout',
    slideMaster: 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideMaster',
    theme: 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme',
    image: 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/image',
    hyperlink: 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink',
    chart: 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/chart',
    presProps: 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/presProps',
    viewProps: 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/viewProps',
    tableStyles: 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/tableStyles'
};

function createElementContext(options = {}) {
    return {
        rels: [],
        media: [],
        nextElementId: 2,
        nextRelId: 2,
        mediaIndex: options.startMediaIndex || 0,
        charts: [],
        chartIndex: 0
    };
}
function addRelationship(ctx, type, target, external) {
    const relId = `rId${ctx.nextRelId++}`;
    ctx.rels.push({ relId, type, target, external: !!external });
    return relId;
}
async function resolveImageData(el) {
    if (el.data) {
        const str = String(el.data);
        const dataUrlMatch = str.match(/^data:image\/([a-z0-9.+-]+);base64,(.+)$/i);
        if (dataUrlMatch) {
            let ext = dataUrlMatch[1].toLowerCase();
            if (ext === 'jpg')
                ext = 'jpeg';
            if (ext === 'svg+xml')
                ext = 'svg';
            return { base64: dataUrlMatch[2], ext: el.extension || ext };
        }
        return { base64: str, ext: el.extension || 'png' };
    }
    if (el.src) {
        if (typeof fetch !== 'function') {
            throw new Error(`无法获取远程图片 ${el.src}：当前环境不支持 fetch。请将图片下载后以 data(base64/dataURL) 方式提供`);
        }
        const resp = await fetch(el.src);
        if (!resp.ok) {
            throw new Error(`下载远程图片失败: ${el.src} (HTTP ${resp.status})`);
        }
        const buffer = await resp.arrayBuffer();
        const base64 = arrayBufferToBase64(buffer);
        const mimeFromHeader = (resp.headers.get('content-type') || '').match(/^image\/([a-z0-9.+-]+)/i);
        let ext = mimeFromHeader ? mimeFromHeader[1].toLowerCase() : null;
        if (ext === 'jpg')
            ext = 'jpeg';
        if (ext === 'svg+xml')
            ext = 'svg';
        return { base64, ext: el.extension || ext || guessExtFromUrl(el.src) || 'png' };
    }
    throw new Error("图片元素需要提供 data（dataURL/base64）或 src（远程 URL）字段");
}
function guessExtFromUrl(url) {
    const match = String(url).split('?')[0].match(/\.([a-z0-9]+)$/i);
    return match ? match[1].toLowerCase() : null;
}
function arrayBufferToBase64(buffer) {
    const bytes = new Uint8Array(buffer);
    let binary = '';
    const chunk = 0x8000;
    for (let i = 0; i < bytes.length; i += chunk) {
        binary += String.fromCharCode.apply(null, bytes.subarray(i, i + chunk));
    }
    if (typeof btoa === 'function') {
        return btoa(binary);
    }
    return Buffer.from(bytes).toString('base64');
}
function buildXfrm(el) {
    return xmlNode('a:xfrm', { rot: el.rotation ? degToRot(el.rotation) : null }, xmlNode('a:off', { x: pxToEmu(el.x || 0), y: pxToEmu(el.y || 0) }), xmlNode('a:ext', { cx: pxToEmu(el.width || 0), cy: pxToEmu(el.height || 0) }));
}
function buildHyperlink(ctx, href) {
    if (!href)
        return null;
    const internalMatch = String(href).match(/^#(\d+)$/);
    if (internalMatch) {
        const relId = addRelationship(ctx, REL_TYPES.slide, `slide${internalMatch[1]}.xml`);
        return xmlNode('a:hlinkClick', { 'r:id': relId, action: 'ppaction://hlinksldjump' });
    }
    const relId = addRelationship(ctx, REL_TYPES.hyperlink, String(href), true);
    return xmlNode('a:hlinkClick', { 'r:id': relId });
}
function buildTextRun(ctx, text, opts = {}) {
    const rPrChildren = [];
    if (opts.color) {
        rPrChildren.push(xmlNode('a:solidFill', xmlNode('a:srgbClr', { val: colorToHex(opts.color) })));
    }
    if (opts.fontFace) {
        rPrChildren.push(xmlNode('a:latin', { typeface: opts.fontFace }));
    }
    const hlink = buildHyperlink(ctx, opts.href);
    if (hlink)
        rPrChildren.push(hlink);
    return xmlNode('a:r', null, xmlNode('a:rPr', {
        lang: opts.lang || 'zh-CN',
        sz: opts.fontSize !== undefined ? ptToSz(opts.fontSize) : null,
        b: opts.bold ? 1 : null,
        i: opts.italic ? 1 : null,
        u: opts.underline ? 'sng' : null,
        dirty: 0
    }, ...rPrChildren), xmlNode('a:t', null, String(text)));
}
function buildParagraph(ctx, paragraph, defaults) {
    const alignMap = { left: 'l', center: 'ctr', right: 'r', justify: 'just' };
    const p = paragraph || {};
    const pPrChildren = [];
    if (p.bullet) {
        pPrChildren.push(xmlNode('a:buFont', { typeface: 'Arial' }));
        pPrChildren.push(xmlNode('a:buChar', { char: '•' }));
    }
    else {
        pPrChildren.push(xmlNode('a:buNone'));
    }
    const pPr = xmlNode('a:pPr', { algn: alignMap[p.align || defaults.align] || null }, ...pPrChildren);
    let runs;
    if (Array.isArray(p.runs) && p.runs.length > 0) {
        runs = p.runs.map((r) => {
            const { text, options, ...flat } = r;
            return buildTextRun(ctx, text, { ...defaults, ...(options || {}), ...flat });
        });
    }
    else {
        runs = [buildTextRun(ctx, p.text !== undefined ? p.text : '', defaults)];
    }
    return xmlNode('a:p', null, pPr, ...runs);
}
function normalizeParagraphs(el) {
    if (Array.isArray(el.paragraphs) && el.paragraphs.length > 0) {
        return el.paragraphs;
    }
    if (Array.isArray(el.runs) && el.runs.length > 0) {
        return [{ runs: el.runs }];
    }
    if (el.text !== undefined) {
        return String(el.text).split('\n').map(t => ({ text: t }));
    }
    return [{ text: '' }];
}
function buildTextElement(ctx, el) {
    const id = ctx.nextElementId++;
    const anchorMap = { top: null, middle: 'ctr', bottom: 'b' };
    const defaults = {
        align: el.align,
        fontSize: el.fontSize,
        color: el.color,
        bold: el.bold,
        italic: el.italic,
        underline: el.underline,
        fontFace: el.fontFace,
        href: el.href,
        lang: el.lang
    };
    return xmlNode('p:sp', null, xmlNode('p:nvSpPr', null, xmlNode('p:cNvPr', { id, name: el.name || `TextBox ${id - 1}` }), xmlNode('p:cNvSpPr', { txBox: 1 }), xmlNode('p:nvPr')), xmlNode('p:spPr', null, buildXfrm(el), xmlNode('a:prstGeom', { prst: 'rect' }, xmlNode('a:avLst'))), xmlNode('p:txBody', null, xmlNode('a:bodyPr', { wrap: 'square', rtlCol: 0, anchor: anchorMap[el.valign] || null }), xmlNode('a:lstStyle'), ...normalizeParagraphs(el).map((p) => buildParagraph(ctx, p, defaults))));
}
function buildShapeElement(ctx, el) {
    const id = ctx.nextElementId++;
    let fillNode;
    if (el.fill === 'none' || el.fill === null) {
        fillNode = xmlNode('a:noFill');
    }
    else {
        const fillColor = typeof el.fill === 'string' ? el.fill : (el.fill && el.fill.color);
        if (fillColor) {
            fillNode = xmlNode('a:solidFill', xmlNode('a:srgbClr', { val: colorToHex(fillColor) }));
        }
        else {
            fillNode = null;
        }
    }
    let lineNode;
    if (el.line === 'none' || el.line === null) {
        lineNode = xmlNode('a:ln', null, xmlNode('a:noFill'));
    }
    else if (el.line) {
        const w = el.line.width !== undefined ? el.line.width : 1;
        lineNode = xmlNode('a:ln', { w: ptToEmu(w) }, xmlNode('a:solidFill', xmlNode('a:srgbClr', { val: colorToHex(el.line.color) })));
    }
    return xmlNode('p:sp', null, xmlNode('p:nvSpPr', null, xmlNode('p:cNvPr', { id, name: el.name || `Shape ${id - 1}` }), xmlNode('p:cNvSpPr'), xmlNode('p:nvPr')), xmlNode('p:spPr', null, buildXfrm(el), xmlNode('a:prstGeom', { prst: el.shapeType || 'rect' }, xmlNode('a:avLst')), fillNode, lineNode));
}
async function buildImageElement(ctx, el) {
    const id = ctx.nextElementId++;
    const { base64, ext } = await resolveImageData(el);
    ctx.mediaIndex++;
    const mediaName = `image${ctx.mediaIndex}.${ext}`;
    ctx.media.push({ name: mediaName, base64 });
    const embedRelId = addRelationship(ctx, REL_TYPES.image, `../media/${mediaName}`);
    let cNvPrChildren = null;
    if (el.href && !/^#\d+$/.test(String(el.href))) {
        const hlinkRelId = addRelationship(ctx, REL_TYPES.hyperlink, String(el.href), true);
        cNvPrChildren = xmlNode('a:hlinkClick', { 'r:id': hlinkRelId });
    }
    return xmlNode('p:pic', null, xmlNode('p:nvPicPr', null, xmlNode('p:cNvPr', { id, name: el.name || `Image ${id - 1}` }, cNvPrChildren), xmlNode('p:cNvPicPr', null, xmlNode('a:picLocks', { noChangeAspect: 1 })), xmlNode('p:nvPr')), xmlNode('p:blipFill', null, xmlNode('a:blip', { 'r:embed': embedRelId }), xmlNode('a:stretch', null, xmlNode('a:fillRect'))), xmlNode('p:spPr', null, buildXfrm(el), xmlNode('a:prstGeom', { prst: 'rect' }, xmlNode('a:avLst'))));
}
function buildChartElement(ctx, el) {
    const id = ctx.nextElementId++;
    ctx.chartIndex++;
    const chartNum = ctx.chartIndex;
    const chartName = `chart${chartNum}.xml`;
    const relId = addRelationship(ctx, REL_TYPES.chart, `../charts/${chartName}`);
    const xml = buildChartXml(el);
    ctx.charts.push({ name: chartName, xml });
    return xmlNode('p:graphicFrame', null, xmlNode('p:nvGraphicFramePr', null, xmlNode('p:cNvPr', { id, name: el.name || `Chart ${chartNum}` }), xmlNode('p:cNvGraphicFramePr'), xmlNode('p:nvPr')), buildXfrm(el), xmlNode('a:graphic', null, xmlNode('a:graphicData', { uri: NS.c }, xmlNode('c:chart', { 'xmlns:c': NS.c, 'xmlns:r': NS.r, 'r:id': relId }))));
}
function strRefXml(values, col) {
    const n = values.length;
    let pts = '';
    for (let i = 0; i < n; i++) {
        pts += `<c:pt idx="${i}"><c:v>${escapeXml(String(values[i]))}</c:v></c:pt>`;
    }
    const last = n > 0 ? n - 1 : 0;
    return `<c:strRef><c:f>Sheet1!$${col}$2:$${col}$${2 + last}</c:f>` +
        `<c:strCache><c:ptCount val="${n}"/>${pts}</c:strCache></c:strRef>`;
}
function numRefXml(values, col) {
    const n = values.length;
    let pts = '';
    for (let i = 0; i < n; i++) {
        pts += `<c:pt idx="${i}"><c:v>${Number(values[i])}</c:v></c:pt>`;
    }
    const last = n > 0 ? n - 1 : 0;
    return `<c:numRef><c:f>Sheet1!$${col}$2:$${col}$${2 + last}</c:f>` +
        `<c:numCache><c:fmtCode>General</c:fmtCode><c:ptCount val="${n}"/>${pts}</c:numCache></c:numRef>`;
}
function buildChartXml(el) {
    const type = el.chartType || 'barChart';
    const isPie = /pie/i.test(type);
    const isScatter = type === 'scatterChart';
    const cats = el.categories || [];
    const series = el.series || [];
    const varyColors = el.varyColors !== undefined ? (el.varyColors ? 1 : 0) : (isPie ? 1 : 0);
    const serXml = series.map((s, i) => {
        const tx = `<c:tx><c:strRef><c:f>Sheet1!$A$1</c:f>` +
            `<c:strCache><c:ptCount val="1"/><c:pt idx="0"><c:v>${escapeXml(s.name || `Series${i + 1}`)}</c:v></c:pt></c:strCache></c:strRef></c:tx>`;
        let data;
        if (isScatter) {
            data = `<c:xVal>${numRefXml(s.x || [], 'B')}</c:xVal><c:yVal>${numRefXml(s.y || [], 'C')}</c:yVal>`;
        }
        else {
            data = `<c:cat>${strRefXml(cats, 'A')}</c:cat><c:val>${numRefXml(s.values || [], 'B')}</c:val>`;
        }
        return `<c:ser><c:idx val="${i}"/><c:order val="${i}"/>${tx}${data}</c:ser>`;
    }).join('');
    let plotChart;
    if (isPie) {
        plotChart = `<c:${type}><c:varyColors val="${varyColors}"/>${serXml}</c:${type}>`;
    }
    else {
        const dir = type === 'barChart' ? `<c:barDir val="${el.barDir || 'col'}"/>` : '';
        const grouping = type === 'lineChart' ? '<c:grouping val="standard"/>' : '';
        plotChart = `<c:${type}>${dir}${grouping}<c:varyColors val="${varyColors}"/>${serXml}` +
            `<c:axId val="111"/><c:axId val="112"/></c:${type}>`;
    }
    let axes = '';
    if (!isPie) {
        axes = `<c:catAx><c:axId val="111"/><c:scaling><c:orientation val="minMax"/></c:scaling><c:delete val="0"/><c:axPos val="b"/><c:crossAx val="112"/></c:catAx><c:valAx><c:axId val="112"/><c:scaling><c:orientation val="minMax"/></c:scaling><c:delete val="0"/><c:axPos val="l"/><c:crossAx val="111"/><c:majorGridlines/></c:valAx>`;
    }
    const titleXml = el.title
        ? `<c:title><c:tx><c:rich><a:bodyPr/><a:lstStyle/>` +
            `<a:p><a:r><a:rPr lang="zh-CN"/><a:t>${escapeXml(el.title)}</a:t></a:r></a:p>` +
            `</c:rich></c:tx><c:overlay val="0"/></c:title>`
        : '';
    const legendXml = el.legend !== false
        ? '<c:legend><c:legendPos val="r"/><c:overlay val="0"/></c:legend>'
        : '';
    const autoTitleDeleted = `<c:autoTitleDeleted val="${el.title ? 0 : 1}"/>`;
    return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\r\n` +
        `<c:chartSpace xmlns:c="${NS.c}" xmlns:a="${NS.a}" xmlns:r="${NS.r}">` +
        `<c:chart>${titleXml}${autoTitleDeleted}` +
        `<c:plotArea><c:layout/>${plotChart}${axes}</c:plotArea>` +
        `${legendXml}<c:plotVisOnly val="1"/></c:chart></c:chartSpace>`;
}
async function buildElement(ctx, el) {
    if (!el || typeof el !== 'object')
        return null;
    switch (el.type) {
        case 'text':
            return buildTextElement(ctx, el);
        case 'shape':
            return buildShapeElement(ctx, el);
        case 'image':
            return buildImageElement(ctx, el);
        case 'chart':
            return buildChartElement(ctx, el);
        default:
            return null;
    }
}
async function buildSlideRoot(ctx, slide) {
    const elementNodes = [];
    for (const el of (slide && slide.elements) || []) {
        const node = await buildElement(ctx, el);
        if (node)
            elementNodes.push(node);
    }
    let bgNode = null;
    const bgColor = slide && slide.background;
    if (bgColor && bgColor !== 'none') {
        bgNode = xmlNode('p:bg', null, xmlNode('p:bgPr', null, xmlNode('a:solidFill', xmlNode('a:srgbClr', { val: colorToHex(bgColor) })), xmlNode('a:effectLst')));
    }
    return xmlNode('p:sld', { 'xmlns:a': 'http://schemas.openxmlformats.org/drawingml/2006/main',
        'xmlns:r': 'http://schemas.openxmlformats.org/officeDocument/2006/relationships',
        'xmlns:p': 'http://schemas.openxmlformats.org/presentationml/2006/main' }, xmlNode('p:cSld', null, bgNode, xmlNode('p:spTree', null, xmlNode('p:nvGrpSpPr', null, xmlNode('p:cNvPr', { id: 1, name: '' }), xmlNode('p:cNvGrpSpPr'), xmlNode('p:nvPr')), xmlNode('p:grpSpPr', null, xmlNode('a:xfrm', null, xmlNode('a:off', { x: 0, y: 0 }), xmlNode('a:ext', { cx: 0, cy: 0 }), xmlNode('a:chOff', { x: 0, y: 0 }), xmlNode('a:chExt', { cx: 0, cy: 0 }))), ...elementNodes)), xmlNode('p:clrMapOvr', null, xmlNode('a:masterClrMapping')));
}

function normalizePresentation(input) {
    let presentation = input;
    if (presentation && typeof presentation.toJSON === 'function') {
        presentation = presentation.toJSON();
    }
    if (!presentation || typeof presentation !== 'object') {
        throw new Error('jsonToPptx: 输入必须为演示文稿 JSON 对象或 PPTXComposer 实例');
    }
    if (!Array.isArray(presentation.slides) || presentation.slides.length === 0) {
        throw new Error('jsonToPptx: 演示文稿至少需要一页幻灯片（slides 数组为空）');
    }
    return presentation;
}
async function jsonToPptx(presentation, options = {}) {
    const pres = normalizePresentation(presentation);
    const slideSize = pres.slideSize || { width: 1280, height: 720 };
    const zip = new JSZip();
    const slideRefs = [];
    const allMediaExts = new Set();
    const allChartNames = [];
    let presRelId = 2;
    for (const i of pres.slides.keys()) {
        const slideIndex = i + 1;
        const ctx = createElementContext();
        const slideRoot = await buildSlideRoot(ctx, pres.slides[i]);
        zip.file(`ppt/slides/slide${slideIndex}.xml`, toXmlDocument(slideRoot));
        const slideRels = [
            { relId: 'rId1', type: REL_TYPES.slideLayout, target: '../slideLayouts/slideLayout1.xml' },
            ...ctx.rels
        ];
        zip.file(`ppt/slides/_rels/slide${slideIndex}.xml.rels`, buildRelationshipsXml(slideRels));
        for (const media of ctx.media) {
            zip.file(`ppt/media/${media.name}`, media.base64, { base64: true });
            allMediaExts.add(media.name.split('.').pop().toLowerCase());
        }
        for (const chart of ctx.charts) {
            zip.file(`ppt/charts/${chart.name}`, chart.xml);
            allChartNames.push(chart.name);
        }
        slideRefs.push({ relId: `rId${presRelId++}`, target: `slides/slide${slideIndex}.xml` });
    }
    zip.file('ppt/presentation.xml', buildPresentationXml(slideSize, slideRefs));
    const presRels = [
        { relId: 'rId1', type: REL_TYPES.slideMaster, target: 'slideMasters/slideMaster1.xml' },
        ...slideRefs.map(ref => ({ relId: ref.relId, type: REL_TYPES.slide, target: ref.target }))
    ];
    presRels.push({ relId: `rId${presRelId++}`, type: REL_TYPES.theme, target: 'theme/theme1.xml' }, { relId: `rId${presRelId++}`, type: REL_TYPES.presProps, target: 'presProps.xml' }, { relId: `rId${presRelId++}`, type: REL_TYPES.viewProps, target: 'viewProps.xml' }, { relId: `rId${presRelId++}`, type: REL_TYPES.tableStyles, target: 'tableStyles.xml' });
    zip.file('ppt/_rels/presentation.xml.rels', buildRelationshipsXml(presRels));
    zip.file('ppt/theme/theme1.xml', buildThemeXml());
    zip.file('ppt/slideMasters/slideMaster1.xml', buildSlideMasterXml());
    zip.file('ppt/slideMasters/_rels/slideMaster1.xml.rels', buildRelationshipsXml(MASTER_RELS));
    zip.file('ppt/slideLayouts/slideLayout1.xml', buildSlideLayoutXml());
    zip.file('ppt/slideLayouts/_rels/slideLayout1.xml.rels', buildRelationshipsXml(LAYOUT_RELS));
    zip.file('ppt/presProps.xml', buildPresPropsXml());
    zip.file('ppt/viewProps.xml', buildViewPropsXml());
    zip.file('ppt/tableStyles.xml', buildTableStylesXml());
    zip.file('docProps/core.xml', buildCorePropsXml(pres.metadata));
    zip.file('docProps/app.xml', buildAppPropsXml(pres.slides.length));
    zip.file('_rels/.rels', buildRootRelsXml());
    let contentTypeXml = buildContentTypesXml([...allMediaExts], pres.slides.length);
    for (const chartName of allChartNames) {
        const override = `<Override PartName="/ppt/charts/${chartName}" ContentType="application/vnd.openxmlformats-officedocument.drawingml.chart+xml"/>`;
        contentTypeXml = contentTypeXml.replace('</Types>', `${override}</Types>`);
    }
    zip.file('[Content_Types].xml', contentTypeXml);
    return zip.generateAsync({
        type: options.outputType || 'uint8array',
        compression: 'DEFLATE',
        compressionOptions: { level: 6 }
    });
}
function parseRelationships(relsText) {
    const rels = [];
    const re = /<Relationship\s+([^>]*?)\/>/g;
    const attrRe = /(\w+)="([^"]*)"/g;
    let match;
    while ((match = re.exec(relsText)) !== null) {
        const attrs = {};
        let attrMatch;
        while ((attrMatch = attrRe.exec(match[1])) !== null) {
            attrs[attrMatch[1]] = attrMatch[2];
        }
        if (attrs.Id)
            rels.push(attrs);
    }
    return rels;
}
function parseSldIdLst(presentationText) {
    const lstMatch = presentationText.match(/<p:sldIdLst>([\s\S]*?)<\/p:sldIdLst>/);
    if (!lstMatch)
        return [];
    const entries = [];
    const re = /<p:sldId\s+([^>]*?)\/>/g;
    const attrRe = /([\w:.-]+)="([^"]*)"/g;
    let match;
    while ((match = re.exec(lstMatch[1])) !== null) {
        const attrs = {};
        let attrMatch;
        while ((attrMatch = attrRe.exec(match[1])) !== null) {
            attrs[attrMatch[1]] = attrMatch[2];
        }
        if (attrs['r:id'])
            entries.push({ id: attrs.id, relId: attrs['r:id'] });
    }
    return entries;
}
function rewriteSldIdLst(presentationText, entries) {
    const inner = entries
        .map((e, i) => `<p:sldId id="${e.id || 256 + i}" r:id="${escapeXml(e.relId)}"/>`)
        .join('');
    return presentationText.replace(/<p:sldIdLst>[\s\S]*?<\/p:sldIdLst>/, `<p:sldIdLst>${inner}</p:sldIdLst>`);
}
function listSlideNumbers(zip) {
    const numbers = [];
    zip.forEach((path, entry) => {
        if (entry.dir)
            return;
        const m = path.match(/^ppt\/slides\/slide(\d+)\.xml$/);
        if (m)
            numbers.push(Number(m[1]));
    });
    return numbers.sort((a, b) => a - b);
}
function maxMediaIndex(zip) {
    let max = 0;
    zip.forEach((path, entry) => {
        if (entry.dir)
            return;
        const m = path.match(/^ppt\/media\/[^/]*?(\d+)\.[^/]+$/);
        if (m)
            max = Math.max(max, Number(m[1]));
    });
    return max;
}
async function editPptx(fileData) {
    const zip = await JSZip.loadAsync(fileData);
    const readText = async (name) => {
        const f = zip.file(name);
        return f ? await f.async('string') : null;
    };
    async function getPresentationInfo() {
        const text = await readText('ppt/presentation.xml');
        if (!text)
            throw new Error('editPptx: 无效的 PPTX 文件（缺少 ppt/presentation.xml）');
        const relsText = await readText('ppt/_rels/presentation.xml.rels');
        const rels = parseRelationships(relsText || '');
        const relById = {};
        for (const r of rels)
            relById[r.Id] = r;
        const entries = parseSldIdLst(text);
        const slideParts = entries.map(e => {
            const rel = relById[e.relId];
            if (!rel)
                return null;
            const target = rel.Target.replace(/^\//, '').replace(/^ppt\//, 'ppt/');
            return { relId: e.relId, target: target.startsWith('ppt/') ? target : `ppt/${target}` };
        }).filter(Boolean);
        return { text, rels, relById, entries, slideParts };
    }
    async function setAppSlideCount(n) {
        const app = await readText('docProps/app.xml');
        if (app && /<Slides>\d+<\/Slides>/.test(app)) {
            zip.file('docProps/app.xml', app.replace(/<Slides>\d+<\/Slides>/, `<Slides>${n}</Slides>`));
        }
    }
    async function save(options = {}) {
        return zip.generateAsync({
            type: options.outputType || 'uint8array',
            compression: 'DEFLATE',
            compressionOptions: { level: 6 }
        });
    }
    return {
        zip,
        save,
        async getSlideCount() {
            const info = await getPresentationInfo();
            return info.entries.length;
        },
        async getSlide(slideNum) {
            const info = await getPresentationInfo();
            const part = info.slideParts[slideNum - 1];
            if (!part)
                throw new Error(`getSlide: 页码越界（共 ${info.entries.length} 页）`);
            return PPTXXmlUtils.readXmlFile(zip, part.target);
        },
        async deleteSlide(slideNum) {
            const info = await getPresentationInfo();
            if (slideNum < 1 || slideNum > info.entries.length) {
                throw new Error(`deleteSlide: 页码越界（共 ${info.entries.length} 页）`);
            }
            if (info.entries.length <= 1) {
                throw new Error('deleteSlide: 至少需保留一页幻灯片，无法删除最后一页');
            }
            const part = info.slideParts[slideNum - 1];
            const slidePath = part.target;
            const remaining = info.entries.filter((_, i) => i !== slideNum - 1);
            zip.file('ppt/presentation.xml', rewriteSldIdLst(info.text, remaining));
            const relsText = await readText('ppt/_rels/presentation.xml.rels');
            const relRe = new RegExp(`<Relationship\\s+Id="${part.relId}"[^>]*/>`);
            zip.file('ppt/_rels/presentation.xml.rels', relsText.replace(relRe, ''));
            const slideRelsText = await readText(`${slidePath.replace('slides/', 'slides/_rels/')}.rels`);
            if (slideRelsText) {
                for (const rel of parseRelationships(slideRelsText)) {
                    if (rel.Type && rel.Type.endsWith('/notesSlide')) {
                        const notesPath = rel.Target.replace('../', 'ppt/');
                        zip.remove(notesPath);
                        zip.remove(`${notesPath.replace('notesSlides/', 'notesSlides/_rels/')}.rels`);
                        await removeContentTypeOverride(notesPath);
                    }
                }
            }
            zip.remove(slidePath);
            zip.remove(`${slidePath.replace('slides/', 'slides/_rels/')}.rels`);
            await removeContentTypeOverride(slidePath);
            await setAppSlideCount(remaining.length);
        },
        async moveSlide(from, to) {
            const info = await getPresentationInfo();
            const entries = [...info.entries];
            if (from < 1 || from > entries.length || to < 1 || to > entries.length) {
                throw new Error(`moveSlide: 页码越界（共 ${entries.length} 页）`);
            }
            const [moved] = entries.splice(from - 1, 1);
            entries.splice(to - 1, 0, moved);
            const orderedRels = entries.map(e => info.relById[e.relId]);
            const oldNums = orderedRels.map(rel => {
                const m = rel.Target.match(/slide(\d+)\.xml$/);
                if (!m)
                    throw new Error(`moveSlide: 无法解析 slide 目标 ${rel.Target}`);
                return Number(m[1]);
            });
            const mapping = {};
            oldNums.forEach((oldNum, i) => { mapping[oldNum] = i + 1; });
            const contents = [];
            const relsContents = [];
            for (const oldNum of oldNums) {
                contents.push(await readText(`ppt/slides/slide${oldNum}.xml`));
                relsContents.push(await readText(`ppt/slides/_rels/slide${oldNum}.xml.rels`));
            }
            oldNums.forEach((oldNum, i) => {
                const newNum = i + 1;
                const newRels = (relsContents[i] || '').replace(/Target="slide(\d+)\.xml"/g, (m, n) => `Target="slide${mapping[n] !== undefined ? mapping[n] : n}.xml"`);
                zip.file(`ppt/slides/slide${newNum}.xml`, contents[i]);
                zip.file(`ppt/slides/_rels/slide${newNum}.xml.rels`, newRels);
            });
            let relsText = await readText('ppt/_rels/presentation.xml.rels');
            orderedRels.forEach((rel, i) => {
                const oldNum = oldNums[i];
                const newNum = i + 1;
                if (oldNum !== newNum) {
                    const re = new RegExp(`(<Relationship\\s+Id="${rel.Id}"[^>]*?Target=")slides/slide${oldNum}\\.xml(")`);
                    relsText = relsText.replace(re, `$1slides/slide${newNum}.xml$2`);
                }
            });
            zip.file('ppt/_rels/presentation.xml.rels', relsText);
            zip.file('ppt/presentation.xml', rewriteSldIdLst(info.text, entries));
        },
        async setMetadata(metadata) {
            const existing = await readText('docProps/core.xml');
            if (existing) {
                zip.file('docProps/core.xml', buildCorePropsXml(metadata));
            }
            else {
                zip.file('docProps/core.xml', buildCorePropsXml(metadata));
                const rootRels = await readText('_rels/.rels');
                if (rootRels && !rootRels.includes('core-properties')) {
                    const newRel = '<Relationship Id="rIdCore" Type="http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties" Target="docProps/core.xml"/>';
                    zip.file('_rels/.rels', rootRels.replace('</Relationships>', `${newRel}</Relationships>`));
                }
                const ctText = await readText('[Content_Types].xml');
                if (ctText && !ctText.includes('core-properties')) {
                    const override = '<Override PartName="/docProps/core.xml" ContentType="application/vnd.openxmlformats-package.core-properties+xml"/>';
                    zip.file('[Content_Types].xml', ctText.replace('</Types>', `${override}</Types>`));
                }
            }
        },
        async addSlide(slideJson) {
            const info = await getPresentationInfo();
            const numbers = listSlideNumbers(zip);
            const nextNum = (numbers.length ? numbers[numbers.length - 1] : 0) + 1;
            const ctx = createElementContext({ startMediaIndex: maxMediaIndex(zip) });
            const slideRoot = await buildSlideRoot(ctx, slideJson);
            zip.file(`ppt/slides/slide${nextNum}.xml`, toXmlDocument(slideRoot));
            const slideRels = [
                { relId: 'rId1', type: REL_TYPES.slideLayout, target: '../slideLayouts/slideLayout1.xml' },
                ...ctx.rels
            ];
            zip.file(`ppt/slides/_rels/slide${nextNum}.xml.rels`, buildRelationshipsXml(slideRels));
            const mediaExts = new Set();
            const newCtCharts = [];
            for (const media of ctx.media) {
                zip.file(`ppt/media/${media.name}`, media.base64, { base64: true });
                mediaExts.add(media.name.split('.').pop().toLowerCase());
            }
            for (const chart of ctx.charts) {
                zip.file(`ppt/charts/${chart.name}`, chart.xml);
                newCtCharts.push(chart.name);
            }
            const maxId = info.entries.reduce((m, e) => Math.max(m, Number(e.id) || 256), 255);
            const usedRelIds = new Set(info.rels.map(r => r.Id));
            let relNum = 1;
            while (usedRelIds.has(`rId${relNum}`))
                relNum++;
            const newRelId = `rId${relNum}`;
            const newEntries = [...info.entries, { id: maxId + 1, relId: newRelId }];
            zip.file('ppt/presentation.xml', rewriteSldIdLst(info.text, newEntries));
            await setAppSlideCount(newEntries.length);
            const relsText = await readText('ppt/_rels/presentation.xml.rels');
            const newRel = `<Relationship Id="${newRelId}" Type="${REL_TYPES.slide}" Target="slides/slide${nextNum}.xml"/>`;
            zip.file('ppt/_rels/presentation.xml.rels', relsText.replace('</Relationships>', `${newRel}</Relationships>`));
            const ctText = await readText('[Content_Types].xml');
            let newCt = ctText;
            const override = `<Override PartName="/ppt/slides/slide${nextNum}.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slide+xml"/>`;
            newCt = newCt.replace('</Types>', `${override}</Types>`);
            const MIME_MAP = { png: 'image/png', jpeg: 'image/jpeg', jpg: 'image/jpeg', gif: 'image/gif', bmp: 'image/bmp', svg: 'image/svg+xml' };
            for (const ext of mediaExts) {
                if (!newCt.includes(`Extension="${ext}"`)) {
                    newCt = newCt.replace('</Types>', `<Default Extension="${ext}" ContentType="${MIME_MAP[ext] || 'application/octet-stream'}"/></Types>`);
                }
            }
            for (const chartName of newCtCharts) {
                if (!newCt.includes(`/ppt/charts/${chartName}`)) {
                    newCt = newCt.replace('</Types>', `<Override PartName="/ppt/charts/${chartName}" ContentType="application/vnd.openxmlformats-officedocument.drawingml.chart+xml"/></Types>`);
                }
            }
            zip.file('[Content_Types].xml', newCt);
        }
    };
    async function removeContentTypeOverride(partPath) {
        const ctText = await readText('[Content_Types].xml');
        if (!ctText)
            return;
        const re = new RegExp(`<Override PartName="/${partPath.replace(/\//g, '\\/')}"[^>]*/>`);
        zip.file('[Content_Types].xml', ctText.replace(re, ''));
    }
}

const DEFAULT_SLIDE_SIZE = { width: 1280, height: 720 };
function makeFluent(el, keys) {
    const builder = {};
    for (const key of keys) {
        builder[key] = (value) => {
            el[key] = value === undefined ? true : value;
            return builder;
        };
    }
    return builder;
}
function applyConfig(element, config) {
    if (typeof config === 'function') {
        config(element);
    }
    else if (config && typeof config === 'object') {
        Object.assign(element, config);
    }
}
class SlideComposer {
    constructor() {
        this.slide = { background: null, elements: [] };
    }
    background(color) {
        this.slide.background = color;
        return this;
    }
    addText(config) {
        const el = { type: 'text', x: 0, y: 0, width: 300, height: 60 };
        if (typeof config === 'function') {
            const builder = makeFluent(el, ['x', 'y', 'width', 'height', 'align', 'valign',
                'fontSize', 'color', 'bold', 'italic', 'underline', 'fontFace', 'href', 'lang', 'name']);
            builder.value = (text) => { el.text = text; return builder; };
            builder.runs = (runs) => { el.runs = runs; return builder; };
            builder.paragraphs = (paragraphs) => { el.paragraphs = paragraphs; return builder; };
            config(builder);
        }
        else {
            applyConfig(el, config);
        }
        this.slide.elements.push(el);
        return this;
    }
    addShape(config) {
        const el = { type: 'shape', shapeType: 'rect', x: 0, y: 0, width: 200, height: 120 };
        if (typeof config === 'function') {
            const builder = makeFluent(el, ['x', 'y', 'width', 'height', 'rotation', 'name']);
            builder.shapeType = (type) => { el.shapeType = type; return builder; };
            builder.fill = (fill) => { el.fill = fill; return builder; };
            builder.line = (line) => { el.line = line; return builder; };
            config(builder);
        }
        else {
            applyConfig(el, config);
        }
        this.slide.elements.push(el);
        return this;
    }
    addImage(config) {
        const el = { type: 'image', x: 0, y: 0, width: 300, height: 200 };
        if (typeof config === 'function') {
            const builder = makeFluent(el, ['x', 'y', 'width', 'height', 'extension', 'href', 'name']);
            builder.data = (data) => { el.data = data; return builder; };
            builder.src = (src) => { el.src = src; return builder; };
            config(builder);
        }
        else {
            applyConfig(el, config);
        }
        this.slide.elements.push(el);
        return this;
    }
    addChart(config) {
        const el = { type: 'chart', chartType: 'barChart', x: 0, y: 0, width: 600, height: 400 };
        if (typeof config === 'function') {
            config(el);
        }
        else {
            applyConfig(el, config);
        }
        this.slide.elements.push(el);
        return this;
    }
}
class PPTXComposer {
    constructor() {
        this.presentation = {
            metadata: {},
            slideSize: { ...DEFAULT_SLIDE_SIZE },
            slides: []
        };
    }
    slideSize(width, height) {
        if (width && typeof width === 'object') {
            this.presentation.slideSize = {
                width: width.width,
                height: width.height
            };
        }
        else {
            this.presentation.slideSize = { width, height };
        }
        return this;
    }
    metadata(metadata) {
        this.presentation.metadata = { ...this.presentation.metadata, ...metadata };
        return this;
    }
    title(value) { return this.metadata({ title: value }); }
    author(value) { return this.metadata({ author: value }); }
    subject(value) { return this.metadata({ subject: value }); }
    keywords(value) { return this.metadata({ keywords: value }); }
    description(value) { return this.metadata({ description: value }); }
    addSlide(config) {
        const slideComposer = new SlideComposer();
        if (typeof config === 'function') {
            config(slideComposer);
        }
        else if (config && typeof config === 'object') {
            applyConfig(slideComposer.slide, config);
        }
        this.presentation.slides.push(slideComposer.slide);
        return this;
    }
    toJSON() {
        return JSON.parse(JSON.stringify(this.presentation));
    }
    save(options) {
        return jsonToPptx(this.toJSON(), options);
    }
}

async function processToJson(file, settings, callbacks, chartId, styleTable, defaultTextStyle) {
    if (file.byteLength < 10) {
        if (callbacks.onError) {
            callbacks.onError({ type: "file_error", message: "Invalid file: file too small" });
        }
        throw new Error("Invalid file: file too small");
    }
    const msgQueue = [];
    const zip = JSZip.loadAsync ? await JSZip.loadAsync(file) : new JSZip().load(file);
    const parsedData = await parsePPTXInternal(zip, msgQueue, settings, chartId, styleTable, defaultTextStyle);
    return {
        parsedData,
        msgQueue,
        zip,
        slideSize: parsedData.slideSize,
        thumbnail: parsedData.thumbnail,
        metadata: parsedData.metadata,
        executionTime: parsedData.executionTime
    };
}
async function parsePPTXInternal(zip, msgQueue, settings, chartId, styleTable, defaultTextStyle) {
    const dateBefore = new Date();
    const thumbFile = zip.file("docProps/thumbnail.jpeg");
    let thumbnail = null;
    if (thumbFile !== null) {
        const pptxThumbImg = PPTXXmlUtils.base64ArrayBuffer(await thumbFile.async("arraybuffer"));
        thumbnail = pptxThumbImg;
    }
    let metadata = {};
    try {
        const coreFile = zip.file("docProps/core.xml");
        if (coreFile !== null) {
            const coreContent = await PPTXXmlUtils.readXmlFile(zip, "docProps/core.xml");
            if (coreContent !== null) {
                const coreProperties = coreContent["cp:coreProperties"];
                if (coreProperties) {
                    metadata = {
                        title: coreProperties["dc:title"] || undefined,
                        subject: coreProperties["dc:subject"] || undefined,
                        author: coreProperties["dc:creator"] || undefined,
                        keywords: coreProperties["cp:keywords"] || undefined,
                        description: coreProperties["dc:description"] || undefined,
                        lastModifiedBy: coreProperties["cp:lastModifiedBy"] || undefined,
                        created: coreProperties["dcterms:created"] || undefined,
                        modified: coreProperties["dcterms:modified"] || undefined,
                        category: coreProperties["cp:category"] || undefined,
                        status: coreProperties["cp:contentStatus"] || undefined,
                        contentType: coreProperties["dc:type"] || undefined,
                        language: coreProperties["dc:language"] || undefined
                    };
                }
            }
        }
    }
    catch (error) {
        metadata = {};
    }
    const filesInfo = await PPTXXmlUtils.getContentTypes(zip);
    const slideSize = await PPTXXmlUtils.getSlideSizeAndSetDefaultTextStyle(zip, settings);
    const slides = [];
    const numOfSlides = filesInfo.slides.length;
    for (let i = 0; i < numOfSlides; i++) {
        const filename = filesInfo.slides[i];
        let fileNameNoPath = "";
        if (filename.includes("/")) {
            const pathParts = filename.split("/");
            fileNameNoPath = pathParts.pop();
        }
        else {
            fileNameNoPath = filename;
        }
        let fileNameNoExt = "";
        if (fileNameNoPath.includes(".")) {
            const nameParts = fileNameNoPath.split(".");
            nameParts.pop();
            fileNameNoExt = nameParts.join(".");
        }
        let slideNumber = 1;
        if (fileNameNoExt !== "" && fileNameNoPath.includes("slide")) {
            slideNumber = Number(fileNameNoExt.substring(5));
        }
        const slideData = await processSingleSlideStructured(zip, filename, i, slideSize, msgQueue, settings, chartId, styleTable, defaultTextStyle);
        slides.push({
            slideNum: slideNumber,
            fileName: fileNameNoExt,
            data: slideData
        });
    }
    slides.sort((a, b) => a.slideNum - b.slideNum);
    const dateAfter = new Date();
    return {
        slides,
        slideSize,
        thumbnail,
        metadata,
        executionTime: dateAfter.getTime() - dateBefore.getTime()
    };
}
async function processSingleSlideStructured(zip, slideFileName, index, slideSize, msgQueue, settings, chartId, styleTable, defaultTextStyle) {
    const resName = `${slideFileName.replace("slides/slide", "slides/_rels/slide")}.rels`;
    const resContent = await PPTXXmlUtils.readXmlFile(zip, resName);
    const relationshipArray = resContent.Relationships.Relationship;
    let layoutFilename = "";
    let diagramFilename = "";
    let notesFilename = "";
    const slideResObj = {};
    if (Array.isArray(relationshipArray)) {
        for (const rel of relationshipArray) {
            const relType = rel.attrs.Type;
            const target = rel.attrs.Target.replace("../", "ppt/").replace(/^\/+/, "");
            switch (relType) {
                case "http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideLayout":
                    layoutFilename = target;
                    break;
                case "http://schemas.microsoft.com/office/2007/relationships/diagramDrawing":
                    diagramFilename = target;
                    slideResObj[rel.attrs.Id] = {
                        type: relType.replace("http://schemas.openxmlformats.org/officeDocument/2006/relationships/", ""),
                        target
                    };
                    break;
                case "http://schemas.openxmlformats.org/officeDocument/2006/relationships/notesSlide":
                    notesFilename = target;
                    break;
                default:
                    slideResObj[rel.attrs.Id] = {
                        type: relType.replace("http://schemas.openxmlformats.org/officeDocument/2006/relationships/", ""),
                        target
                    };
            }
        }
    }
    else {
        const relType = relationshipArray.attrs.Type;
        const target = relationshipArray.attrs.Target.replace("../", "ppt/").replace(/^\/+/, "");
        if (relType === "http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideLayout") {
            layoutFilename = target;
        }
        else if (relType === "http://schemas.openxmlformats.org/officeDocument/2006/relationships/notesSlide") {
            notesFilename = target;
        }
        else if (relType === "http://schemas.microsoft.com/office/2007/relationships/diagramDrawing") {
            diagramFilename = target;
            slideResObj[relationshipArray.attrs.Id] = {
                type: relType.replace("http://schemas.openxmlformats.org/officeDocument/2006/relationships/", ""),
                target: target
            };
        }
        else {
            layoutFilename = target;
        }
    }
    const slideLayoutContent = await PPTXXmlUtils.readXmlFile(zip, layoutFilename);
    const slideLayoutTables = PPTXNodeUtils.indexNodes(slideLayoutContent);
    const layoutColorOverride = PPTXXmlUtils.getTextByPathList(slideLayoutContent, ["p:sldLayout", "p:clrMapOvr", "a:overrideClrMapping"]);
    if (layoutColorOverride !== undefined) {
        layoutColorOverride.attrs;
    }
    const slideLayoutResFilename = `${layoutFilename.replace("slideLayouts/slideLayout", "slideLayouts/_rels/slideLayout")}.rels`;
    const slideLayoutResContent = await PPTXXmlUtils.readXmlFile(zip, slideLayoutResFilename);
    const layoutRelArray = slideLayoutResContent.Relationships.Relationship;
    let masterFilename = "";
    const layoutResObj = {};
    if (Array.isArray(layoutRelArray)) {
        for (const rel of layoutRelArray) {
            const relType = rel.attrs.Type;
            const target = rel.attrs.Target.replace("../", "ppt/").replace(/^\/+/, "");
            if (relType === "http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideMaster") {
                masterFilename = target;
            }
            else {
                layoutResObj[rel.attrs.Id] = {
                    type: relType.replace("http://schemas.openxmlformats.org/officeDocument/2006/relationships/", ""),
                    target
                };
            }
        }
    }
    else {
        masterFilename = layoutRelArray.attrs.Target.replace("../", "ppt/").replace(/^\/+/, "");
    }
    const slideMasterContent = await PPTXXmlUtils.readXmlFile(zip, masterFilename);
    const slideMasterTextStyles = PPTXXmlUtils.getTextByPathList(slideMasterContent, ["p:sldMaster", "p:txStyles"]);
    const slideMasterTables = PPTXNodeUtils.indexNodes(slideMasterContent);
    const slideMasterResFilename = `${masterFilename.replace("slideMasters/slideMaster", "slideMasters/_rels/slideMaster")}.rels`;
    const slideMasterResContent = await PPTXXmlUtils.readXmlFile(zip, slideMasterResFilename);
    const masterRelArray = slideMasterResContent.Relationships.Relationship;
    let themeFilename = "";
    const masterResObj = {};
    if (Array.isArray(masterRelArray)) {
        for (const rel of masterRelArray) {
            const relType = rel.attrs.Type;
            const target = rel.attrs.Target.replace("../", "ppt/").replace(/^\/+/, "");
            if (relType === "http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme") {
                themeFilename = target;
            }
            else {
                masterResObj[rel.attrs.Id] = {
                    type: relType.replace("http://schemas.openxmlformats.org/officeDocument/2006/relationships/", ""),
                    target
                };
            }
        }
    }
    else {
        themeFilename = masterRelArray.attrs.Target.replace("../", "ppt/").replace(/^\/+/, "");
    }
    let themeContent;
    const themeResObj = {};
    if (themeFilename !== undefined) {
        const themeName = themeFilename.split("/").pop();
        const themeResFileName = `${themeFilename.replace(themeName, `_rels/${themeName}`)}.rels`;
        themeContent = await PPTXXmlUtils.readXmlFile(zip, themeFilename);
        const themeResContent = await PPTXXmlUtils.readXmlFile(zip, themeResFileName);
        if (themeResContent !== null) {
            const themeRelArray = themeResContent.Relationships.Relationship;
            if (themeRelArray !== undefined) {
                if (Array.isArray(themeRelArray)) {
                    for (const rel of themeRelArray) {
                        themeResObj[rel.attrs.Id] = {
                            type: rel.attrs.Type.replace("http://schemas.openxmlformats.org/officeDocument/2006/relationships/", ""),
                            target: rel.attrs.Target.replace("../", "ppt/").replace(/^\/+/, "")
                        };
                    }
                }
                else {
                    themeResObj[themeRelArray.attrs.Id] = {
                        type: themeRelArray.attrs.Type.replace("http://schemas.openxmlformats.org/officeDocument/2006/relationships/", ""),
                        target: themeRelArray.attrs.Target.replace("../", "ppt/").replace(/^\/+/, "")
                    };
                }
            }
        }
    }
    let diagramContent = {};
    const diagramResObj = {};
    if (diagramFilename !== undefined) {
        const diagramName = diagramFilename.split("/").pop();
        const diagramResFileName = `${diagramFilename.replace(diagramName, `_rels/${diagramName}`)}.rels`;
        diagramContent = await PPTXXmlUtils.readXmlFile(zip, diagramFilename);
        if (diagramContent !== null && diagramContent !== undefined && diagramContent !== "") {
            const diagramJson = JSON.stringify(diagramContent);
            const cleanedJson = diagramJson.replace(/dsp:/g, "p:");
            diagramContent = JSON.parse(cleanedJson);
        }
        const diagramResContent = await PPTXXmlUtils.readXmlFile(zip, diagramResFileName);
        if (diagramResContent !== null) {
            const diagramRelArray = diagramResContent.Relationships.Relationship;
            if (Array.isArray(diagramRelArray)) {
                for (const rel of diagramRelArray) {
                    diagramResObj[rel.attrs.Id] = {
                        type: rel.attrs.Type.replace("http://schemas.openxmlformats.org/officeDocument/2006/relationships/", ""),
                        target: rel.attrs.Target.replace("../", "ppt/").replace(/^\/+/, "")
                    };
                }
            }
            else {
                diagramResObj[diagramRelArray.attrs.Id] = {
                    type: diagramRelArray.attrs.Type.replace("http://schemas.openxmlformats.org/officeDocument/2006/relationships/", ""),
                    target: diagramRelArray.attrs.Target.replace("../", "ppt/").replace(/^\/+/, "")
                };
            }
        }
    }
    const tableStyles = await PPTXXmlUtils.readXmlFile(zip, "ppt/tableStyles.xml");
    const slideContent = await PPTXXmlUtils.readXmlFile(zip, slideFileName, true);
    let notesContent = null;
    if (notesFilename) {
        notesContent = await PPTXXmlUtils.readXmlFile(zip, notesFilename);
    }
    slideContent["p:sld"]["p:cSld"]["p:spTree"];
    settings.themeProcess;
    return {
        slideLayoutContent,
        slideLayoutTables,
        slideMasterContent,
        slideMasterTables,
        slideContent,
        slideResObj,
        slideMasterTextStyles,
        layoutResObj,
        masterResObj,
        themeContent,
        themeResObj,
        diagramContent,
        diagramResObj,
        notesContent,
        defaultTextStyle: slideSize.defaultTextStyle || defaultTextStyle,
        tableStyles,
        styleTable,
        chartId,
        msgQueue,
        bulletCounter: {},
        slideSize,
        index
    };
}
async function convertSlideDataToHtml(slideData, slideSize, settings, zip, slideNum) {
    const warpObj = {
        slideLayoutContent: slideData.slideLayoutContent,
        slideLayoutTables: slideData.slideLayoutTables,
        slideMasterContent: slideData.slideMasterContent,
        slideMasterTables: slideData.slideMasterTables,
        slideContent: slideData.slideContent,
        slideResObj: slideData.slideResObj,
        slideMasterTextStyles: slideData.slideMasterTextStyles,
        layoutResObj: slideData.layoutResObj,
        masterResObj: slideData.masterResObj,
        themeContent: slideData.themeContent,
        themeResObj: slideData.themeResObj,
        diagramContent: slideData.diagramContent,
        diagramResObj: slideData.diagramResObj,
        defaultTextStyle: slideData.defaultTextStyle,
        tableStyles: slideData.tableStyles,
        styleTable: slideData.styleTable,
        chartId: slideData.chartId,
        msgQueue: slideData.msgQueue,
        bulletCounter: slideData.bulletCounter,
        zip: zip
    };
    const processFullTheme = settings.themeProcess;
    let bgResult = "";
    if (processFullTheme === true) {
        bgResult = await PPTXNodeUtils.getBackground(warpObj, slideSize, slideData.index, settings);
    }
    let bgColor = "";
    if (processFullTheme === "colorsAndImageOnly") {
        bgColor = await PPTXStyleUtils.getSlideBackgroundFill(warpObj, slideData.index);
    }
    let transitionClass = "";
    const transitionData = extractSlideTransition(slideData.slideContent);
    if (transitionData) {
        transitionClass = ` data-transition='${JSON.stringify(transitionData)}'`;
    }
    const slideIdAttr = slideNum ? ` id="slide-${slideNum}"` : "";
    let result = `<section class='slide'${slideIdAttr}${transitionClass} style='width:${slideSize.width}px; height:${slideSize.height}px;${bgColor}'>`;
    result += bgResult;
    const nodes = slideData.slideContent["p:sld"]["p:cSld"]["p:spTree"];
    for (const nodeKey in nodes) {
        if (Array.isArray(nodes[nodeKey])) {
            for (const node of nodes[nodeKey]) {
                result += await PPTXNodeUtils.processNodesInSlide(nodeKey, node, nodes, warpObj, "slide", "group", settings);
            }
        }
        else {
            result += await PPTXNodeUtils.processNodesInSlide(nodeKey, nodes[nodeKey], nodes, warpObj, "slide", "group", settings);
        }
    }
    return `${result}</div></section>`;
}
function genGlobalCSS(styleTable) {
    let cssText = "";
    for (const key in styleTable) {
        const suffix = styleTable[key].suffix || "";
        cssText += ` .${styleTable[key].name}${suffix}{${styleTable[key].text}}\n`;
    }
    return cssText;
}
async function pptxToHtml(fileData, options) {
    const settings = {
        mediaProcess: true,
        themeProcess: true,
        incSlide: {
            width: 0,
            height: 0
        },
        styleTable: {},
        ...options
    };
    const callbacks = settings.callbacks || {};
    let defaultTextStyle = null;
    const chartId = { value: 0 };
    const styleTable = settings.styleTable;
    if (callbacks.onFileStart) {
        callbacks.onFileStart();
    }
    async function convertToHtml(file) {
        const { parsedData, msgQueue, zip, slideSize, thumbnail, metadata, executionTime } = await processToJson(file, settings, callbacks, chartId, styleTable, defaultTextStyle);
        const result = {
            slides: [],
            slideSize,
            thumbnail,
            styles: {
                global: ""
            },
            metadata,
            charts: []
        };
        for (const slideData of parsedData.slides) {
            const slideHtml = await convertSlideDataToHtml(slideData.data, slideSize, settings, zip, slideData.slideNum);
            result.slides.push({
                html: slideHtml,
                data: slideData.data,
                slideNum: slideData.slideNum,
                fileName: slideData.fileName
            });
            if (callbacks.onSlide) {
                callbacks.onSlide(slideHtml, {
                    slideNum: slideData.slideNum,
                    fileName: slideData.fileName
                });
            }
        }
        result.styles.global = genGlobalCSS(styleTable);
        if (thumbnail && callbacks.onThumbnail) {
            callbacks.onThumbnail(thumbnail);
        }
        if (slideSize && callbacks.onSlideSize) {
            callbacks.onSlideSize(slideSize);
        }
        if (callbacks.onGlobalCSS) {
            callbacks.onGlobalCSS(result.styles.global);
        }
        processMsgQueue(msgQueue, result);
        if (callbacks.onComplete) {
            callbacks.onComplete({
                executionTime,
                slideWidth: slideSize?.width || 0,
                slideHeight: slideSize?.height || 0,
                styleTable,
                settings
            });
        }
        return result;
    }
    if (fileData) {
        return convertToHtml(fileData);
    }
    return null;
}
async function pptxToJson(fileData, options) {
    const settings = {
        mediaProcess: true,
        themeProcess: true,
        incSlide: {
            width: 0,
            height: 0
        },
        styleTable: {},
        ...options
    };
    const callbacks = settings.callbacks || {};
    let defaultTextStyle = null;
    const chartId = { value: 0 };
    const styleTable = settings.styleTable;
    if (callbacks.onFileStart) {
        callbacks.onFileStart();
    }
    async function convertToJson(file) {
        const { parsedData, msgQueue, slideSize, thumbnail, metadata, executionTime } = await processToJson(file, settings, callbacks, chartId, styleTable, defaultTextStyle);
        const result = {
            slides: [],
            slideSize,
            thumbnail,
            styles: {
                global: genGlobalCSS(styleTable)
            },
            metadata,
            charts: []
        };
        for (const slideData of parsedData.slides) {
            result.slides.push({
                data: slideData.data,
                slideNum: slideData.slideNum,
                fileName: slideData.fileName
            });
            if (callbacks.onSlide) {
                callbacks.onSlide(slideData.data, {
                    slideNum: slideData.slideNum,
                    fileName: slideData.fileName
                });
            }
        }
        if (thumbnail && callbacks.onThumbnail) {
            callbacks.onThumbnail(thumbnail);
        }
        if (slideSize && callbacks.onSlideSize) {
            callbacks.onSlideSize(slideSize);
        }
        if (callbacks.onGlobalCSS) {
            callbacks.onGlobalCSS(result.styles.global);
        }
        processMsgQueue(msgQueue, result);
        if (callbacks.onComplete) {
            callbacks.onComplete({
                executionTime,
                slideWidth: slideSize?.width || 0,
                slideHeight: slideSize?.height || 0,
                styleTable,
                settings
            });
        }
        return result;
    }
    if (fileData) {
        return convertToJson(fileData);
    }
    return null;
}
async function pptxToFiles(fileData) {
    if (fileData.byteLength < 10) {
        throw new Error("Invalid file: file too small");
    }
    const zip = JSZip.loadAsync ? await JSZip.loadAsync(fileData) : new JSZip().load(fileData);
    const result = {
        files: [],
        content: {}
    };
    const promises = [];
    zip.forEach((relativePath, zipEntry) => {
        result.files.push({
            name: relativePath,
            dir: zipEntry.dir,
            size: zipEntry._data.uncompressedSize
        });
        const promise = (async () => {
            try {
                if (zipEntry.dir) {
                    return;
                }
                const ext = relativePath.split('.').pop().toLowerCase();
                if (ext === 'xml' || ext === 'rels') {
                    const content = await zipEntry.async('text');
                    result.content[relativePath] = {
                        type: 'text',
                        content: content
                    };
                }
                else if (['png', 'jpg', 'jpeg', 'gif', 'bmp', 'svg'].includes(ext)) {
                    const base64 = await zipEntry.async('base64');
                    result.content[relativePath] = {
                        type: 'image',
                        format: ext,
                        base64: base64,
                        dataUrl: `data:image/${ext === 'jpg' ? 'jpeg' : ext};base64,${base64}`
                    };
                }
                else {
                    const base64 = await zipEntry.async('base64');
                    result.content[relativePath] = {
                        type: 'binary',
                        base64: base64
                    };
                }
            }
            catch (error) {
                result.content[relativePath] = {
                    type: 'error',
                    error: error.message
                };
            }
        })();
        promises.push(promise);
    });
    await Promise.all(promises);
    return result;
}
function extractSlideTransition(slideContent) {
    const sld = slideContent["p:sld"];
    if (!sld)
        return null;
    const transition = PPTXXmlUtils.getTextByPathList(sld, ["p:transition"]);
    if (!transition)
        return null;
    let transitionType = "fade";
    let duration = 1000;
    const transitionTypes = [
        "p:blinds", "p:checker", "p:circle", "p:comb",
        "p:cover", "p:dissolve", "p:fade", "p:push",
        "p:random", "p:split", "p:strips", "p:wipe"
    ];
    for (const type of transitionTypes) {
        if (transition[type]) {
            transitionType = type.replace("p:", "");
            break;
        }
    }
    if (transition.attrs && transition.attrs["spd"]) {
        const speedMap = { "1": 500, "2": 1000, "3": 2000 };
        duration = speedMap[transition.attrs["spd"]] || 1000;
    }
    return {
        type: transitionType,
        duration: duration
    };
}

export { PPTXComposer, pptxToHtml as default, editPptx, jsonToPptx, pptxToFiles, pptxToHtml, pptxToJson };
//# sourceMappingURL=ppt-parser.browser.js.map

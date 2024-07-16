/*! For license information please see taskpane.js.LICENSE.txt */
!function(){"use strict";var e,t,n,r,o,a,c,u,s={14385:function(e){e.exports=function(e,t){return t||(t={}),e?(e=String(e.__esModule?e.default:e),t.hash&&(e+=t.hash),t.maybeNeedQuotes&&/[\t\n\f\r "'=<>`]/.test(e)?'"'.concat(e,'"'):e):e}},98362:function(e,t,n){e.exports=n.p+"assets/logo-filled.png"},98388:function(e,t,n){e.exports=n.p+"b116cb09288947d8842f.js"},44329:function(e,t,n){e.exports=n.p+"70c84b971f446e76f6cd.js"},52678:function(e,t,n){e.exports=n.p+"e7a636d7439a38af653d.js"},54058:function(e,t,n){e.exports=n.p+"b0e07be5d4c3f182dc5f.js"},58394:function(e,t,n){e.exports=n.p+"605055c46ab85cca6529.css"}},i={};function l(e){var t=i[e];if(void 0!==t)return t.exports;var n=i[e]={exports:{}};return s[e](n,n.exports,l),n.exports}l.m=s,l.n=function(e){var t=e&&e.__esModule?function(){return e.default}:function(){return e};return l.d(t,{a:t}),t},l.d=function(e,t){for(var n in t)l.o(t,n)&&!l.o(e,n)&&Object.defineProperty(e,n,{enumerable:!0,get:t[n]})},l.g=function(){if("object"==typeof globalThis)return globalThis;try{return this||new Function("return this")()}catch(e){if("object"==typeof window)return window}}(),l.o=function(e,t){return Object.prototype.hasOwnProperty.call(e,t)},function(){var e;l.g.importScripts&&(e=l.g.location+"");var t=l.g.document;if(!e&&t&&(t.currentScript&&(e=t.currentScript.src),!e)){var n=t.getElementsByTagName("script");if(n.length)for(var r=n.length-1;r>-1&&(!e||!/^http(s?):/.test(e));)e=n[r--].src}if(!e)throw new Error("Automatic publicPath is not supported in this browser");e=e.replace(/#.*$/,"").replace(/\?.*$/,"").replace(/\/[^\/]+$/,"/"),l.p=e}(),l.b=document.baseURI||self.location.href,function(){function e(t){return e="function"==typeof Symbol&&"symbol"==typeof Symbol.iterator?function(e){return typeof e}:function(e){return e&&"function"==typeof Symbol&&e.constructor===Symbol&&e!==Symbol.prototype?"symbol":typeof e},e(t)}function t(){t=function(){return r};var n,r={},o=Object.prototype,a=o.hasOwnProperty,c=Object.defineProperty||function(e,t,n){e[t]=n.value},u="function"==typeof Symbol?Symbol:{},s=u.iterator||"@@iterator",i=u.asyncIterator||"@@asyncIterator",l=u.toStringTag||"@@toStringTag";function p(e,t,n){return Object.defineProperty(e,t,{value:n,enumerable:!0,configurable:!0,writable:!0}),e[t]}try{p({},"")}catch(n){p=function(e,t,n){return e[t]=n}}function f(e,t,n,r){var o=t&&t.prototype instanceof w?t:w,a=Object.create(o.prototype),u=new N(r||[]);return c(a,"_invoke",{value:O(e,n,u)}),a}function h(e,t,n){try{return{type:"normal",arg:e.call(t,n)}}catch(e){return{type:"throw",arg:e}}}r.wrap=f;var d="suspendedStart",v="suspendedYield",m="executing",g="completed",y={};function w(){}function x(){}function b(){}var k={};p(k,s,(function(){return this}));var E=Object.getPrototypeOf,I=E&&E(E(P([])));I&&I!==o&&a.call(I,s)&&(k=I);var A=b.prototype=w.prototype=Object.create(k);function S(e){["next","throw","return"].forEach((function(t){p(e,t,(function(e){return this._invoke(t,e)}))}))}function L(t,n){function r(o,c,u,s){var i=h(t[o],t,c);if("throw"!==i.type){var l=i.arg,p=l.value;return p&&"object"==e(p)&&a.call(p,"__await")?n.resolve(p.__await).then((function(e){r("next",e,u,s)}),(function(e){r("throw",e,u,s)})):n.resolve(p).then((function(e){l.value=e,u(l)}),(function(e){return r("throw",e,u,s)}))}s(i.arg)}var o;c(this,"_invoke",{value:function(e,t){function a(){return new n((function(n,o){r(e,t,n,o)}))}return o=o?o.then(a,a):a()}})}function O(e,t,r){var o=d;return function(a,c){if(o===m)throw Error("Generator is already running");if(o===g){if("throw"===a)throw c;return{value:n,done:!0}}for(r.method=a,r.arg=c;;){var u=r.delegate;if(u){var s=R(u,r);if(s){if(s===y)continue;return s}}if("next"===r.method)r.sent=r._sent=r.arg;else if("throw"===r.method){if(o===d)throw o=g,r.arg;r.dispatchException(r.arg)}else"return"===r.method&&r.abrupt("return",r.arg);o=m;var i=h(e,t,r);if("normal"===i.type){if(o=r.done?g:v,i.arg===y)continue;return{value:i.arg,done:r.done}}"throw"===i.type&&(o=g,r.method="throw",r.arg=i.arg)}}}function R(e,t){var r=t.method,o=e.iterator[r];if(o===n)return t.delegate=null,"throw"===r&&e.iterator.return&&(t.method="return",t.arg=n,R(e,t),"throw"===t.method)||"return"!==r&&(t.method="throw",t.arg=new TypeError("The iterator does not provide a '"+r+"' method")),y;var a=h(o,e.iterator,t.arg);if("throw"===a.type)return t.method="throw",t.arg=a.arg,t.delegate=null,y;var c=a.arg;return c?c.done?(t[e.resultName]=c.value,t.next=e.nextLoc,"return"!==t.method&&(t.method="next",t.arg=n),t.delegate=null,y):c:(t.method="throw",t.arg=new TypeError("iterator result is not an object"),t.delegate=null,y)}function j(e){var t={tryLoc:e[0]};1 in e&&(t.catchLoc=e[1]),2 in e&&(t.finallyLoc=e[2],t.afterLoc=e[3]),this.tryEntries.push(t)}function B(e){var t=e.completion||{};t.type="normal",delete t.arg,e.completion=t}function N(e){this.tryEntries=[{tryLoc:"root"}],e.forEach(j,this),this.reset(!0)}function P(t){if(t||""===t){var r=t[s];if(r)return r.call(t);if("function"==typeof t.next)return t;if(!isNaN(t.length)){var o=-1,c=function e(){for(;++o<t.length;)if(a.call(t,o))return e.value=t[o],e.done=!1,e;return e.value=n,e.done=!0,e};return c.next=c}}throw new TypeError(e(t)+" is not iterable")}return x.prototype=b,c(A,"constructor",{value:b,configurable:!0}),c(b,"constructor",{value:x,configurable:!0}),x.displayName=p(b,l,"GeneratorFunction"),r.isGeneratorFunction=function(e){var t="function"==typeof e&&e.constructor;return!!t&&(t===x||"GeneratorFunction"===(t.displayName||t.name))},r.mark=function(e){return Object.setPrototypeOf?Object.setPrototypeOf(e,b):(e.__proto__=b,p(e,l,"GeneratorFunction")),e.prototype=Object.create(A),e},r.awrap=function(e){return{__await:e}},S(L.prototype),p(L.prototype,i,(function(){return this})),r.AsyncIterator=L,r.async=function(e,t,n,o,a){void 0===a&&(a=Promise);var c=new L(f(e,t,n,o),a);return r.isGeneratorFunction(t)?c:c.next().then((function(e){return e.done?e.value:c.next()}))},S(A),p(A,l,"Generator"),p(A,s,(function(){return this})),p(A,"toString",(function(){return"[object Generator]"})),r.keys=function(e){var t=Object(e),n=[];for(var r in t)n.push(r);return n.reverse(),function e(){for(;n.length;){var r=n.pop();if(r in t)return e.value=r,e.done=!1,e}return e.done=!0,e}},r.values=P,N.prototype={constructor:N,reset:function(e){if(this.prev=0,this.next=0,this.sent=this._sent=n,this.done=!1,this.delegate=null,this.method="next",this.arg=n,this.tryEntries.forEach(B),!e)for(var t in this)"t"===t.charAt(0)&&a.call(this,t)&&!isNaN(+t.slice(1))&&(this[t]=n)},stop:function(){this.done=!0;var e=this.tryEntries[0].completion;if("throw"===e.type)throw e.arg;return this.rval},dispatchException:function(e){if(this.done)throw e;var t=this;function r(r,o){return u.type="throw",u.arg=e,t.next=r,o&&(t.method="next",t.arg=n),!!o}for(var o=this.tryEntries.length-1;o>=0;--o){var c=this.tryEntries[o],u=c.completion;if("root"===c.tryLoc)return r("end");if(c.tryLoc<=this.prev){var s=a.call(c,"catchLoc"),i=a.call(c,"finallyLoc");if(s&&i){if(this.prev<c.catchLoc)return r(c.catchLoc,!0);if(this.prev<c.finallyLoc)return r(c.finallyLoc)}else if(s){if(this.prev<c.catchLoc)return r(c.catchLoc,!0)}else{if(!i)throw Error("try statement without catch or finally");if(this.prev<c.finallyLoc)return r(c.finallyLoc)}}}},abrupt:function(e,t){for(var n=this.tryEntries.length-1;n>=0;--n){var r=this.tryEntries[n];if(r.tryLoc<=this.prev&&a.call(r,"finallyLoc")&&this.prev<r.finallyLoc){var o=r;break}}o&&("break"===e||"continue"===e)&&o.tryLoc<=t&&t<=o.finallyLoc&&(o=null);var c=o?o.completion:{};return c.type=e,c.arg=t,o?(this.method="next",this.next=o.finallyLoc,y):this.complete(c)},complete:function(e,t){if("throw"===e.type)throw e.arg;return"break"===e.type||"continue"===e.type?this.next=e.arg:"return"===e.type?(this.rval=this.arg=e.arg,this.method="return",this.next="end"):"normal"===e.type&&t&&(this.next=t),y},finish:function(e){for(var t=this.tryEntries.length-1;t>=0;--t){var n=this.tryEntries[t];if(n.finallyLoc===e)return this.complete(n.completion,n.afterLoc),B(n),y}},catch:function(e){for(var t=this.tryEntries.length-1;t>=0;--t){var n=this.tryEntries[t];if(n.tryLoc===e){var r=n.completion;if("throw"===r.type){var o=r.arg;B(n)}return o}}throw Error("illegal catch attempt")},delegateYield:function(e,t,r){return this.delegate={iterator:P(e),resultName:t,nextLoc:r},"next"===this.method&&(this.arg=n),y}},r}function n(e,t){return function(e){if(Array.isArray(e))return e}(e)||function(e,t){var n=null==e?null:"undefined"!=typeof Symbol&&e[Symbol.iterator]||e["@@iterator"];if(null!=n){var r,o,a,c,u=[],s=!0,i=!1;try{if(a=(n=n.call(e)).next,0===t){if(Object(n)!==n)return;s=!1}else for(;!(s=(r=a.call(n)).done)&&(u.push(r.value),u.length!==t);s=!0);}catch(e){i=!0,o=e}finally{try{if(!s&&null!=n.return&&(c=n.return(),Object(c)!==c))return}finally{if(i)throw o}}return u}}(e,t)||function(e,t){if(e){if("string"==typeof e)return a(e,t);var n={}.toString.call(e).slice(8,-1);return"Object"===n&&e.constructor&&(n=e.constructor.name),"Map"===n||"Set"===n?Array.from(e):"Arguments"===n||/^(?:Ui|I)nt(?:8|16|32)(?:Clamped)?Array$/.test(n)?a(e,t):void 0}}(e,t)||function(){throw new TypeError("Invalid attempt to destructure non-iterable instance.\nIn order to be iterable, non-array objects must have a [Symbol.iterator]() method.")}()}function r(e,t,n,r,o,a,c){try{var u=e[a](c),s=u.value}catch(e){return void n(e)}u.done?t(s):Promise.resolve(s).then(r,o)}function o(e){return function(){var t=this,n=arguments;return new Promise((function(o,a){var c=e.apply(t,n);function u(e){r(c,o,a,u,s,"next",e)}function s(e){r(c,o,a,u,s,"throw",e)}u(void 0)}))}}function a(e,t){(null==t||t>e.length)&&(t=e.length);for(var n=0,r=Array(t);n<t;n++)r[n]=e[n];return r}var c,u,s,i,l=/[^:0-9]/g,p=new FileReader,f=new DOMParser,h=new XMLSerializer,d={types:[{description:"Json Settings",accept:{"file/*":[".json"]}}],excludeAcceptAllOption:!0,multiple:!1},v={letters:"ABCDEFGHIJKLMNOPQRSTUVWXYZ",numbers:"0123456789"};function m(){return g.apply(this,arguments)}function g(){return(g=o(t().mark((function e(){var r,o,a,c;return t().wrap((function(e){for(;;)switch(e.prev=e.next){case 0:if(e.prev=0,"web"!=s){e.next=9;break}(r=document.createElement("input")).type="file",r.accept=".json",r.click(),r.addEventListener("change",(function(e){var t=e.target.files[0],n=new FileReader;n.onload=function(e){y(JSON.parse(e.target.result))},n.readAsText(t)})),e.next=24;break;case 9:return e.next=11,window.showOpenFilePicker(d);case 11:return o=e.sent,a=n(o,1),i=a[0],e.next=16,i.getFile();case 16:return c=e.sent,e.t0=y,e.t1=JSON,e.next=21,c.text();case 21:e.t2=e.sent,e.t3=e.t1.parse.call(e.t1,e.t2),(0,e.t0)(e.t3);case 24:e.next=29;break;case 26:e.prev=26,e.t4=e.catch(0),console.log(e.t4);case 29:case"end":return e.stop()}}),e,null,[[0,26]])})))).apply(this,arguments)}function y(e){return w.apply(this,arguments)}function w(){return w=o(t().mark((function e(n){return t().wrap((function(e){for(;;)switch(e.prev=e.next){case 0:return e.prev=0,e.next=3,Excel.run(function(){var e=o(t().mark((function e(r){var o,a,s,i,l,p;return t().wrap((function(e){for(;;)switch(e.prev=e.next){case 0:for(o=document.getElementById("excel-trunk"),u=n.fqdn,document.getElementById("excel-fqdn").value=u,a=0;a<n.sbc.length;a++)c.template.sbc.push(n.sbc[a]);for(s=[],i=0;i<n.trunks.length;i++)(l=document.createElement("option")).value=n.trunks[i].number,l.innerHTML=n.trunks[i].name,o.appendChild(l),p=[n.trunks[i].id,n.trunks[i].name,"#123","123456"],s.push(p);r.workbook.worksheets.getItem("Numeri Brevi").getRange("A2:D".concat(s.length+1)).values=s;case 9:case"end":return e.stop()}}),e)})));return function(t){return e.apply(this,arguments)}}());case 3:e.next=8;break;case 5:e.prev=5,e.t0=e.catch(0),console.log(e.t0);case 8:case"end":return e.stop()}}),e,null,[[0,5]])}))),w.apply(this,arguments)}function x(){return b.apply(this,arguments)}function b(){return b=o(t().mark((function e(){return t().wrap((function(e){for(;;)switch(e.prev=e.next){case 0:return e.prev=0,e.next=3,Excel.run(function(){var e=o(t().mark((function e(n){var r,o,a,c,u,s;return t().wrap((function(e){for(;;)switch(e.prev=e.next){case 0:return r=n.workbook.worksheets.getItem("Numeri brevi"),e.next=3,F(n,{columns:{start:"a",end:"d"}});case 3:return o=e.sent,a=r.getRange(o),c=r.getRange("A1:D1"),a.load("values"),c.load("values"),e.next=10,n.sync();case 10:u=[c.values[0]].concat(a.values),s=new Blob([u.join("\n")],{type:"text/plain"}),saveAs(s,"Numeri Brevi.csv");case 13:case"end":return e.stop()}}),e)})));return function(t){return e.apply(this,arguments)}}());case 3:e.next=8;break;case 5:e.prev=5,e.t0=e.catch(0),console.log(e.t0);case 8:case"end":return e.stop()}}),e,null,[[0,5]])}))),b.apply(this,arguments)}function k(){return E.apply(this,arguments)}function E(){return E=o(t().mark((function e(){return t().wrap((function(e){for(;;)switch(e.prev=e.next){case 0:return e.prev=0,e.next=3,Excel.run(function(){var e=o(t().mark((function e(n){var r,o,a,c,u,s;return t().wrap((function(e){for(;;)switch(e.prev=e.next){case 0:return r=n.workbook.worksheets.getItem("GNR"),e.next=3,F(n,{columns:{start:"a",end:"n"}});case 3:return o=e.sent,a=r.getRange(o),c=r.getRange("A1:N1"),a.load("values"),c.load("values"),e.next=10,n.sync();case 10:u=[c.values[0]].concat(a.values),s=new Blob([u.join("\n")],{type:"text/plain"}),saveAs(s,"Gnr.csv");case 13:case"end":return e.stop()}}),e)})));return function(t){return e.apply(this,arguments)}}());case 3:e.next=8;break;case 5:e.prev=5,e.t0=e.catch(0),console.log(e.t0);case 8:case"end":return e.stop()}}),e,null,[[0,5]])}))),E.apply(this,arguments)}function I(){return A.apply(this,arguments)}function A(){return A=o(t().mark((function e(){return t().wrap((function(e){for(;;)switch(e.prev=e.next){case 0:return e.prev=0,e.next=3,Excel.run(function(){var e=o(t().mark((function e(n){var r,o,a,c,u,s;return t().wrap((function(e){for(;;)switch(e.prev=e.next){case 0:return r=n.workbook.worksheets.getItem("Uscita interni"),e.next=3,F(n,{columns:{start:"a",end:"bh"}});case 3:return o=e.sent,a=r.getRange(o),c=r.getRange("A1:BH1"),a.load("values"),c.load("values"),e.next=10,n.sync();case 10:u=[c.values[0]].concat(a.values),s=new Blob([u.join("\n")],{type:"text/plain"}),saveAs(s,"Interni.csv");case 13:case"end":return e.stop()}}),e)})));return function(t){return e.apply(this,arguments)}}());case 3:e.next=8;break;case 5:e.prev=5,e.t0=e.catch(0),console.log(e.t0);case 8:case"end":return e.stop()}}),e,null,[[0,5]])}))),A.apply(this,arguments)}function S(){return L.apply(this,arguments)}function L(){return L=o(t().mark((function e(){return t().wrap((function(e){for(;;)switch(e.prev=e.next){case 0:return e.prev=0,e.next=3,Excel.run(function(){var e=o(t().mark((function e(n){var r,o,a,c,u,s;return t().wrap((function(e){for(;;)switch(e.prev=e.next){case 0:return r=n.workbook.worksheets.getItem("Uscita contatti"),e.next=3,F(n,{columns:{start:"a",end:"n"}});case 3:return o=e.sent,a=r.getRange(o),(c=extOutPage.getRange("A1:BH1")).load("values"),a.load("values"),e.next=10,n.sync();case 10:u=[c.values[0]].concat(a.values),s=new Blob([u.values.join("\n")],{type:"text/plain"}),saveAs(s,"Contatti.csv");case 13:case"end":return e.stop()}}),e)})));return function(t){return e.apply(this,arguments)}}());case 3:e.next=8;break;case 5:e.prev=5,e.t0=e.catch(0),console.log(e.t0);case 8:case"end":return e.stop()}}),e,null,[[0,5]])}))),L.apply(this,arguments)}function O(){return R.apply(this,arguments)}function R(){return R=o(t().mark((function e(){return t().wrap((function(e){for(;;)switch(e.prev=e.next){case 0:return e.prev=0,e.next=3,Excel.run(function(){var e=o(t().mark((function e(n){var r,o,a,s,i,l,p,f,h;return t().wrap((function(e){for(;;)switch(e.prev=e.next){case 0:return u=document.getElementById("excel-fqdn").value,r=n.workbook.worksheets.getItem("Interni"),e.next=4,F(n,{columns:{start:"a",end:"j"}});case 4:return o=e.sent,(a=r.getRange(o)).load("values"),e.next=9,n.sync();case 9:for(s=new Array(a.values.length),i=0;i<a.values.length;i++){for(l=c.template.ext.slice(),l=H(a.values[i],l,c.template.extOffset,c.template.extOffset.length),p=0;p<c.template.pinOffset.length;p++)l[c.template.pinOffset[p]]=D(0,c.credentials.pin.length);for(f=0;f<c.template.passOffset.length;f++)l[c.template.passOffset[f]]=D(1,c.credentials.password.length,c.credentials.password.pattern);""!=a.values[i][6]&&(h=q(a.values[i]),l=H(h,l,c.template.phoneOffset)),console.log(l),s[i]=l}console.log(s),n.workbook.worksheets.getItem("Uscita interni").getRange("A2:BH".concat(s.length+1)).values=s;case 15:case"end":return e.stop()}}),e)})));return function(t){return e.apply(this,arguments)}}());case 3:e.next=8;break;case 5:e.prev=5,e.t0=e.catch(0),console.log(e.t0);case 8:case"end":return e.stop()}}),e,null,[[0,5]])}))),R.apply(this,arguments)}function j(){return B.apply(this,arguments)}function B(){return B=o(t().mark((function e(){return t().wrap((function(e){for(;;)switch(e.prev=e.next){case 0:return e.prev=0,e.next=3,Excel.run(function(){var e=o(t().mark((function e(n){var r,o,a,u,s,i;return t().wrap((function(e){for(;;)switch(e.prev=e.next){case 0:return r=n.workbook.worksheets.getItem("Contatti"),e.next=3,F(n,{columns:{start:"a",end:"i"}});case 3:return o=e.sent,(a=r.getRange(o)).load("values"),e.next=8,n.sync();case 8:for(u=[],s=0;s<a.values.length;s++)i=c.template.phBook.slice(),u.push(H(a.values[s],i,c.template.pbookOffset));n.workbook.worksheets.getItem("Uscita contatti").getRange("A2:N".concat(u.length+1)).values=u;case 13:case"end":return e.stop()}}),e)})));return function(t){return e.apply(this,arguments)}}());case 3:e.next=8;break;case 5:e.prev=5,e.t0=e.catch(0),console.log(e.t0);case 8:case"end":return e.stop()}}),e,null,[[0,5]])}))),B.apply(this,arguments)}function N(){return P.apply(this,arguments)}function P(){return P=o(t().mark((function e(){return t().wrap((function(e){for(;;)switch(e.prev=e.next){case 0:return e.prev=0,e.next=3,Excel.run(function(){var e=o(t().mark((function e(n){var r,o,a,c,u,s,i,l,p,f,h,d,v,m,g,y,w;return t().wrap((function(e){for(;;)switch(e.prev=e.next){case 0:for(r=n.workbook.worksheets.getItem("Interni"),o=n.workbook.worksheets.getItem("GNR"),a=document.getElementById("excel-gnr-number").value,c=document.getElementById("excel-gnr-day").value,u=document.getElementById("excel-gnr-night").value,s=document.getElementById("excel-trunk").value,i=[],l=a.split("-"),p=l[1].length,f=l[0].slice(0,l[0].length-p),l[0]=Number(l[0].slice(-p)),l[1]=Number(l[1]),h=l[0];h<=l[1]-l[0];h++)d=f,2==p&&h<10&&(d+="0"),3==p&&(h<10&&(d+="00"),h>=10&&h<100&&(d+="0")),d+=h,i.push([h-l[0]+2,"GNR",1,d,s,2,c,0,0,,,2,u,0]);return e.next=15,F(n,{columns:{start:"a",end:"a"}});case 15:return v=e.sent,m=r.getRange(v),g=m.load("values"),e.next=20,n.sync();case 20:for(y=0;y<g.values.length;y++)for(w=0;w<i.length;w++)i[w][3].slice(-g.values[y][0].length)==g.values[y][0]&&(i[w][1]="DIRETTO",i[w][6]=g.values[y][0],i[w][12]=g.values[y][0]);o.getRange("A2:N".concat(i.length+1)).values=i;case 23:case"end":return e.stop()}}),e)})));return function(t){return e.apply(this,arguments)}}());case 3:e.next=8;break;case 5:e.prev=5,e.t0=e.catch(0),console.log(e.t0);case 8:case"end":return e.stop()}}),e,null,[[0,5]])}))),P.apply(this,arguments)}function T(){return C.apply(this,arguments)}function C(){return C=o(t().mark((function e(){return t().wrap((function(e){for(;;)switch(e.prev=e.next){case 0:return e.prev=0,e.next=3,Excel.run(function(){var e=o(t().mark((function e(n){var r,o,a,u,s,i,l,p,f;return t().wrap((function(e){for(;;)switch(e.prev=e.next){case 0:return r=n.workbook.worksheets.getItem("Interni"),o=document.getElementById("excel-phone").value,e.next=4,F(n,{columns:{start:"g",end:"g"}});case 4:return a=e.sent,e.next=7,F(n,{columns:{start:"i",end:"i"}});case 7:for(u=e.sent,s=r.getRange(a),i=r.getRange(u),s.values="Sel. Modello",i.values="Sel. SBC",s.dataValidation.rule={list:{inCellDropDown:!0,source:"".concat(c.phones[o].models.join(","))}},l=[],p=0;p<c.template.sbc.length;p++)f=c.template.sbc[p].name,l.push(f);i.dataValidation.rule={list:{inCellDropDown:!0,source:"".concat(l.join(","))}};case 16:case"end":return e.stop()}}),e)})));return function(t){return e.apply(this,arguments)}}());case 3:e.next=8;break;case 5:e.prev=5,e.t0=e.catch(0),console.log(e.t0);case 8:case"end":return e.stop()}}),e,null,[[0,5]])}))),C.apply(this,arguments)}function M(){return _.apply(this,arguments)}function _(){return _=o(t().mark((function e(){return t().wrap((function(e){for(;;)switch(e.prev=e.next){case 0:return e.prev=0,e.next=3,Excel.run(function(){var e=o(t().mark((function e(n){var r,o;return t().wrap((function(e){for(;;)switch(e.prev=e.next){case 0:return r=n.workbook.worksheets.load("items/name"),o=new Array,e.next=4,n.sync();case 4:r.items.forEach((function(e){o.push(e.name)})),c.pages.forEach((function(e,t){0==o.includes(e.name)&&n.workbook.worksheets.add(e.name)}));case 6:case"end":return e.stop()}}),e)})));return function(t){return e.apply(this,arguments)}}());case 3:return e.next=5,Excel.run(function(){var e=o(t().mark((function e(n){return t().wrap((function(e){for(;;)switch(e.prev=e.next){case 0:c.pages.forEach((function(e){var t,r,o,a=(t=e.cells.length-1,r="",(o=Math.trunc(t/26))>0?(r+=String.fromCharCode(65+o-1),r+=String.fromCharCode(65+t%26)):r+=String.fromCharCode(65+t%26),r),c=n.workbook.worksheets.getItem(e.name),u=c.getRange("A1:".concat(a,"1"));u.load("values"),u.values=[e.cells],c.getRange("A:".concat(a)).numberFormat="@"}));case 1:case"end":return e.stop()}}),e)})));return function(t){return e.apply(this,arguments)}}());case 5:e.next=10;break;case 7:e.prev=7,e.t0=e.catch(0),console.log(e.t0);case 10:case"end":return e.stop()}}),e,null,[[0,7]])}))),_.apply(this,arguments)}function F(e,t){return G.apply(this,arguments)}function G(){return(G=o(t().mark((function e(n,r){var o,a,c,u;return t().wrap((function(e){for(;;)switch(e.prev=e.next){case 0:return(o=n.workbook.getSelectedRange()).load("address"),e.next=4,n.sync();case 4:for(a=o.address.split("!")[1],c=a.split(":"),u=0;u<c.length;u++)c[u]=c[u].replace(l,""),1==c[u]&&(c[u]=2);if(1!=c.length||r.columns.start!=r.columns.end){e.next=9;break}return e.abrupt("return","".concat(r.columns.start).concat(c[0]));case 9:if(1!=c.length||r.columns.start==r.columns.end){e.next=12;break}return e.abrupt("return","".concat(r.columns.start).concat(c[0],":").concat(r.columns.end).concat(c[0]));case 12:return e.abrupt("return","".concat(r.columns.start).concat(c[0],":").concat(r.columns.end).concat(c[1]));case 14:case"end":return e.stop()}}),e)})))).apply(this,arguments)}function U(){return(U=o(t().mark((function e(){return t().wrap((function(e){for(;;)switch(e.prev=e.next){case 0:fetch("https://raw.githubusercontent.com/ZSirag/3CX-Excel/main/settings.json").then((function(e){return e.json()})).then((function(e){c=e;for(var t=document.getElementById("excel-phone"),n=0;n<c.phones.length;n++){var r=document.createElement("option");r.innerHTML=c.phones[n].name,r.value=n,t.appendChild(r)}}));case 1:case"end":return e.stop()}}),e)})))).apply(this,arguments)}function D(e,t,n){for(var r=v.numbers,o="",a=0,c=r.length;a<t;++a)o+=r.charAt(Math.floor(Math.random()*c));if(e){o=o.split(""),(n=n.split("")).sort((function(e,t){return.5-Math.random()}));for(var u=0;u<t;u++)"x"==n[u]&&(o[u]=v.letters.charAt(Math.floor(Math.random()*v.letters.length)).toLocaleLowerCase()),"X"==n[u]&&(o[u]=v.letters.charAt(Math.floor(Math.random()*v.letters.length))),"1"==n[u]&&(o[u]=v.numbers.charAt(Math.floor(Math.random()*v.numbers.length)));o=o.join("")}return o}function H(e,t,n,r){if(r)for(var o=0;o<r;o++)t[n[o]]=e[o];else for(var a=0;a<e.length;a++)t[n[a]]=e[a];return t}function q(e){for(var t=[],n=0;n<c.phones.length;n++)if(c.phones[n].models.includes(e[6])){t.push(e[6]),t.push(e[7]),t.push(c.phones[n].xml),t.push(J(c.phones[n].nosbc,e[8],e[9])),t.push(u);break}return t}function J(e,t,n){var r=f.parseFromString(e,"text/xml"),o=r.querySelector("PhoneDevice");if("no"!=t)if("Tel. Router"==t)o.setAttribute("ProvType","3"),o.setAttribute("SbcName",D(1,c.credentials.SBC.length,c.credentials.SBC.pattern)),o.setAttribute("RemoteSpmPort","0"),o.setAttribute("IsSbc","1");else for(var a=0;a<c.template.sbc.length;a++)c.template.sbc[a].name==t&&(o.setAttribute("ProvType","3"),o.setAttribute("SbcName",c.template.sbc[a].id),o.setAttribute("RemoteSpmPort","5060"),o.setAttribute("IsSbc","0"));if(""!=n)for(var u=r.querySelectorAll("option"),s=0;s<u.length;s++)"vlanwanport"==u[s].getAttribute("name")&&u[s].setAttribute("value","true"),"vlanwanid"==u[s].getAttribute("name")&&u[s].setAttribute("value",n);return e=h.serializeToString(r).replaceAll('"','""'),'"'.concat(e,'"')}p.onload=function(e){console.log(e.target.result)},Office.onReady((function(e){e.host===Office.HostType.Excel?(document.getElementById("web-body-container").style.display="none",document.getElementById("excel-container").style.display="grid",document.getElementById("excel-container").onclick=M,document.getElementById("excel-f-add-phone").onclick=T,document.getElementById("excel-gen-gnr").onclick=N,document.getElementById("excel-f-gen-pbook").onclick=j,document.getElementById("excel-f-gen-ext").onclick=O,document.getElementById("excel-f-out-ext").onclick=I,document.getElementById("excel-f-out-pbook").onclick=S,document.getElementById("excel-out-gnr").onclick=k,document.getElementById("excel-out-shortn").onclick=x,document.getElementById("excel-f-import-config").onclick=m,function(){U.apply(this,arguments)}(),window!=window.top&&(s="web")):document.getElementById("btn-outbound-rules").onclick=testWeb}))}(),e=l(14385),t=l.n(e),n=new URL(l(98388),l.b),r=new URL(l(44329),l.b),o=new URL(l(52678),l.b),a=new URL(l(54058),l.b),c=new URL(l(58394),l.b),u=new URL(l(98362),l.b),t()(n),t()(r),t()(o),t()(a),t()(c),t()(u)}();
//# sourceMappingURL=taskpane.js.map
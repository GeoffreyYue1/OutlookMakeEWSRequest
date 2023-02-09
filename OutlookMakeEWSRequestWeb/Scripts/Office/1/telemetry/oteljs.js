var oteljs=function(e){var t={};function n(r){if(t[r])return t[r].exports;var i=t[r]={i:r,l:!1,exports:{}};return e[r].call(i.exports,i,i.exports,n),i.l=!0,i.exports}return n.m=e,n.c=t,n.d=function(e,t,r){n.o(e,t)||Object.defineProperty(e,t,{enumerable:!0,get:r})},n.r=function(e){"undefined"!=typeof Symbol&&Symbol.toStringTag&&Object.defineProperty(e,Symbol.toStringTag,{value:"Module"}),Object.defineProperty(e,"__esModule",{value:!0})},n.t=function(e,t){if(1&t&&(e=n(e)),8&t)return e;if(4&t&&"object"==typeof e&&e&&e.__esModule)return e;var r=Object.create(null);if(n.r(r),Object.defineProperty(r,"default",{enumerable:!0,value:e}),2&t&&"string"!=typeof e)for(var i in e)n.d(r,i,function(t){return e[t]}.bind(null,i));return r},n.n=function(e){var t=e&&e.__esModule?function(){return e.default}:function(){return e};return n.d(t,"a",t),t},n.o=function(e,t){return Object.prototype.hasOwnProperty.call(e,t)},n.p="",n(n.s=0)}([function(e,t,n){e.exports=n(1)},function(e,t,n){"use strict";var r,i,o,a,s,u,c,l,f,d,p,v,y;function h(e,t){return{name:e,dataType:r.Boolean,value:t}}function g(e,t){return i.validateInt(t),{name:e,dataType:r.Int64,value:t}}function m(e,t){return{name:e,dataType:r.Double,value:t}}function S(e,t){return{name:e,dataType:r.String,value:t}}function T(e,t,n){var r=n.map((function(t){return{name:e+"."+t.name,value:t.value,dataType:t.dataType}}));return F(r,e,t),r}function F(e,t,n){e.push(S("zC."+t,n))}n.r(t),function(e){e[e.String=0]="String",e[e.Boolean=1]="Boolean",e[e.Int64=2]="Int64",e[e.Double=3]="Double"}(r||(r={})),function(e){var t=-9007199254740991,n=9007199254740991,i=/^[A-Z][a-zA-Z0-9]*$/,o=/^[a-zA-Z0-9_\.]*$/;function a(e){return void 0!==e&&o.test(e)}function s(e){if(!((t=e.name)&&a(t)&&t.length+5<100))throw new Error("Invalid dataField name");var t;e.dataType===r.Int64&&u(e.value)}function u(e){if("number"!=typeof e||!isFinite(e)||Math.floor(e)!==e||e<t||e>n)throw{message:"Invalid integer "+JSON.stringify(e)}}e.validateTelemetryEvent=function(e){if(!function(e){if(!e||e.length>98)return!1;var t=e.split("."),n=t[t.length-1];return function(e){return!!e&&e.length>=3&&"Office"===e[0]}(t)&&(r=n,void 0!==r&&i.test(r));var r}(e.eventName))throw new Error("Invalid eventName");if(e.eventContract&&!a(e.eventContract.name))throw new Error("Invalid eventContract");if(null!=e.dataFields)for(var t=0;t<e.dataFields.length;t++)s(e.dataFields[t])},e.validateInt=u}(i||(i={})),a=o||(o={}),s="Office.System.Result",a.getFields=function(e,t){var n=[];return n.push(g(e+".Code",t.code)),void 0!==t.type&&n.push(S(e+".Type",t.type)),void 0!==t.tag&&n.push(g(e+".Tag",t.tag)),void 0!==t.isExpected&&n.push(h(e+".IsExpected",t.isExpected)),F(n,e,s),n},(c=u||(u={})).contractName="Office.System.Activity",c.getFields=function(e){var t=[];return void 0!==e.cV&&t.push(S("Activity.CV",e.cV)),t.push(g("Activity.Duration",e.duration)),t.push(g("Activity.Count",e.count)),t.push(g("Activity.AggMode",e.aggMode)),void 0!==e.success&&t.push(h("Activity.Success",e.success)),void 0!==e.result&&t.push.apply(t,o.getFields("Activity.Result",e.result)),F(t,"Activity",c.contractName),t},function(e){var t="Office.System.Host";e.getFields=function(e,n){var r=[];return void 0!==n.id&&r.push(S(e+".Id",n.id)),void 0!==n.version&&r.push(S(e+".Version",n.version)),void 0!==n.sessionId&&r.push(S(e+".SessionId",n.sessionId)),F(r,e,t),r}}(l||(l={})),function(e){var t="Office.System.User";e.getFields=function(e,n){var r=[];return void 0!==n.alias&&r.push(S(e+".Alias",n.alias)),void 0!==n.primaryIdentityHash&&r.push(S(e+".PrimaryIdentityHash",n.primaryIdentityHash)),void 0!==n.primaryIdentitySpace&&r.push(S(e+".PrimaryIdentitySpace",n.primaryIdentitySpace)),void 0!==n.tenantId&&r.push(S(e+".TenantId",n.tenantId)),void 0!==n.tenantGroup&&r.push(S(e+".TenantGroup",n.tenantGroup)),void 0!==n.isAnonymous&&r.push(h(e+".IsAnonymous",n.isAnonymous)),F(r,e,t),r}}(f||(f={})),function(e){var t="Office.System.SDX";e.getFields=function(e,n){var r=[];return void 0!==n.id&&r.push(S(e+".Id",n.id)),void 0!==n.version&&r.push(S(e+".Version",n.version)),void 0!==n.instanceId&&r.push(S(e+".InstanceId",n.instanceId)),void 0!==n.name&&r.push(S(e+".Name",n.name)),void 0!==n.marketplaceType&&r.push(S(e+".MarketplaceType",n.marketplaceType)),void 0!==n.sessionId&&r.push(S(e+".SessionId",n.sessionId)),void 0!==n.browserToken&&r.push(S(e+".BrowserToken",n.browserToken)),void 0!==n.osfRuntimeVersion&&r.push(S(e+".OsfRuntimeVersion",n.osfRuntimeVersion)),void 0!==n.officeJsVersion&&r.push(S(e+".OfficeJsVersion",n.officeJsVersion)),void 0!==n.hostJsVersion&&r.push(S(e+".HostJsVersion",n.hostJsVersion)),void 0!==n.assetId&&r.push(S(e+".AssetId",n.assetId)),void 0!==n.providerName&&r.push(S(e+".ProviderName",n.providerName)),void 0!==n.type&&r.push(S(e+".Type",n.type)),F(r,e,t),r}}(d||(d={})),function(e){var t="Office.System.Funnel";e.getFields=function(e,n){var r=[];return void 0!==n.name&&r.push(S(e+".Name",n.name)),void 0!==n.state&&r.push(S(e+".State",n.state)),F(r,e,t),r}}(p||(p={})),function(e){var t="Office.System.UserAction";e.getFields=function(e,n){var r=[];return void 0!==n.id&&r.push(g(e+".Id",n.id)),void 0!==n.name&&r.push(S(e+".Name",n.name)),void 0!==n.commandSurface&&r.push(S(e+".CommandSurface",n.commandSurface)),void 0!==n.parentName&&r.push(S(e+".ParentName",n.parentName)),void 0!==n.triggerMethod&&r.push(S(e+".TriggerMethod",n.triggerMethod)),void 0!==n.timeOffsetMs&&r.push(g(e+".TimeOffsetMs",n.timeOffsetMs)),F(r,e,t),r}}(v||(v={})),function(e){var t="Office.System.Error";e.getFields=function(e,n){var r=[];return r.push(S(e+".ErrorGroup",n.errorGroup)),r.push(g(e+".Tag",n.tag)),void 0!==n.code&&r.push(g(e+".Code",n.code)),void 0!==n.id&&r.push(g(e+".Id",n.id)),void 0!==n.count&&r.push(g(e+".Count",n.count)),F(r,e,t),r}}(y||(y={}));var b,w,N=u,C=o,E=y,A=p,I=l,k=d,x=v,_=f;!function(e){!function(e){!function(e){e.Activity=N,e.Result=C,e.Error=E,e.Funnel=A,e.Host=I,e.SDX=k,e.User=_,e.UserAction=x}(e.System||(e.System={}))}(e.Office||(e.Office={}))}(b||(b={})),function(e){var t,n=0;e.getNext=function(){return void 0===t&&(t=function(){for(var e="ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz0123456789+/",t=[],n=0;n<22;n++)t.push(e.charAt(Math.floor(Math.random()*e.length)));return t.join("")}()),new r(t,++n)},e.getNextChild=function(e){return new r(e.getString(),++e.nextChild)};var r=function(){function e(e,t){this.base=e,this.id=t,this.nextChild=0}return e.prototype.getString=function(){return this.base+"."+this.id},e}();e.CV=r}(w||(w={}));var O,P,D=function(){function e(){this._listeners=[]}return e.prototype.fireEvent=function(e){this._listeners.forEach((function(t){return t(e)}))},e.prototype.addListener=function(e){e&&this._listeners.push(e)},e.prototype.removeListener=function(e){this._listeners=this._listeners.filter((function(t){return t!==e}))},e.prototype.getListenerCount=function(){return this._listeners.length},e}(),M=new D;function j(){return M}function V(e,t,n){M.fireEvent({level:e,category:t,message:n})}!function(e){e[e.Error=0]="Error",e[e.Warning=1]="Warning",e[e.Info=2]="Info",e[e.Verbose=3]="Verbose"}(O||(O={})),function(e){e[e.Core=0]="Core",e[e.Sink=1]="Sink",e[e.Transport=2]="Transport"}(P||(P={}));var L=function(e,t,n,r){return new(n||(n=Promise))((function(i,o){function a(e){try{u(r.next(e))}catch(e){o(e)}}function s(e){try{u(r.throw(e))}catch(e){o(e)}}function u(e){var t;e.done?i(e.value):(t=e.value,t instanceof n?t:new n((function(e){e(t)}))).then(a,s)}u((r=r.apply(e,t||[])).next())}))},U=function(e,t){var n,r,i,o,a={label:0,sent:function(){if(1&i[0])throw i[1];return i[1]},trys:[],ops:[]};return o={next:s(0),throw:s(1),return:s(2)},"function"==typeof Symbol&&(o[Symbol.iterator]=function(){return this}),o;function s(o){return function(s){return function(o){if(n)throw new TypeError("Generator is already executing.");for(;a;)try{if(n=1,r&&(i=2&o[0]?r.return:o[0]?r.throw||((i=r.return)&&i.call(r),0):r.next)&&!(i=i.call(r,o[1])).done)return i;switch(r=0,i&&(o=[2&o[0],i.value]),o[0]){case 0:case 1:i=o;break;case 4:return a.label++,{value:o[1],done:!1};case 5:a.label++,r=o[1],o=[0];continue;case 7:o=a.ops.pop(),a.trys.pop();continue;default:if(!(i=(i=a.trys).length>0&&i[i.length-1])&&(6===o[0]||2===o[0])){a=0;continue}if(3===o[0]&&(!i||o[1]>i[0]&&o[1]<i[3])){a.label=o[1];break}if(6===o[0]&&a.label<i[1]){a.label=i[1],i=o;break}if(i&&a.label<i[2]){a.label=i[2],a.ops.push(o);break}i[2]&&a.ops.pop(),a.trys.pop();continue}o=t.call(e,a)}catch(e){o=[6,e],r=0}finally{n=i=0}if(5&o[0])throw o[1];return{value:o[0]?o[1]:void 0,done:!0}}([o,s])}}},H=function(){return 1e3*Date.now()};"object"==typeof window.performance&&"now"in window.performance&&(H=function(){return 1e3*Math.floor(window.performance.now())});var B,J,R,G,z,W,Z,X,$=function(){function e(e,t,n){this._optionalEventFlags={},this._ended=!1,this._telemetryLogger=e,this._activityName=t,this._cv=n?w.getNextChild(n._cv):w.getNext(),this._dataFields=[],this._success=void 0,this._startTime=H()}return e.createNew=function(t,n){return new e(t,n)},e.prototype.createChildActivity=function(t){return new e(this._telemetryLogger,t,this)},e.prototype.setEventFlags=function(e){this._optionalEventFlags=e},e.prototype.addDataField=function(e){this._dataFields.push(e)},e.prototype.addDataFields=function(e){var t;(t=this._dataFields).push.apply(t,e)},e.prototype.setSuccess=function(e){this._success=e},e.prototype.setResult=function(e,t,n){this._result={code:e,type:t,tag:n}},e.prototype.endNow=function(){if(!this._ended){void 0===this._success&&void 0===this._result&&V(O.Warning,P.Core,(function(){return"Activity does not have success or result set"}));var e=H()-this._startTime;this._ended=!0;var t={duration:e,count:1,aggMode:0,cV:this._cv.getString(),success:this._success,result:this._result};return this._telemetryLogger.sendActivity(this._activityName,t,this._dataFields,this._optionalEventFlags)}V(O.Error,P.Core,(function(){return"Activity has already ended"}))},e.prototype.executeAsync=function(e){return L(this,void 0,void 0,(function(){var t=this;return U(this,(function(n){return[2,e(this).then((function(e){return t.endNow(),e})).catch((function(e){throw t.endNow(),e}))]}))}))},e.prototype.executeSync=function(e){try{var t=e(this);return this.endNow(),t}catch(e){throw this.endNow(),e}},e.prototype.executeChildActivityAsync=function(e,t){return L(this,void 0,void 0,(function(){return U(this,(function(n){return[2,this.createChildActivity(e).executeAsync(t)]}))}))},e.prototype.executeChildActivitySync=function(e,t){return this.createChildActivity(e).executeSync(t)},e}();function q(e){var t={costPriority:G.Normal,samplingPolicy:J.Measure,persistencePriority:R.Normal,dataCategories:z.NotSet,diagnosticLevel:W.FullEvent};return e.eventFlags&&e.eventFlags.dataCategories||V(O.Error,P.Core,(function(){return"Event is missing DataCategories event flag"})),e.eventFlags?(e.eventFlags.costPriority&&(t.costPriority=e.eventFlags.costPriority),e.eventFlags.samplingPolicy&&(t.samplingPolicy=e.eventFlags.samplingPolicy),e.eventFlags.persistencePriority&&(t.persistencePriority=e.eventFlags.persistencePriority),e.eventFlags.dataCategories&&(t.dataCategories=e.eventFlags.dataCategories),e.eventFlags.diagnosticLevel&&(t.diagnosticLevel=e.eventFlags.diagnosticLevel),t):t}!function(e){e[e.EssentialServiceMetadata=1]="EssentialServiceMetadata",e[e.AccountData=2]="AccountData",e[e.SystemMetadata=4]="SystemMetadata",e[e.OrganizationIdentifiableInformation=8]="OrganizationIdentifiableInformation",e[e.EndUserIdentifiableInformation=16]="EndUserIdentifiableInformation",e[e.CustomerContent=32]="CustomerContent",e[e.AccessControl=64]="AccessControl"}(B||(B={})),function(e){e[e.NotSet=0]="NotSet",e[e.Measure=1]="Measure",e[e.Diagnostics=2]="Diagnostics",e[e.CriticalBusinessImpact=191]="CriticalBusinessImpact",e[e.CriticalCensus=192]="CriticalCensus",e[e.CriticalExperimentation=193]="CriticalExperimentation",e[e.CriticalUsage=194]="CriticalUsage"}(J||(J={})),function(e){e[e.NotSet=0]="NotSet",e[e.Normal=1]="Normal",e[e.High=2]="High"}(R||(R={})),function(e){e[e.NotSet=0]="NotSet",e[e.Normal=1]="Normal",e[e.High=2]="High"}(G||(G={})),function(e){e[e.NotSet=0]="NotSet",e[e.SoftwareSetup=1]="SoftwareSetup",e[e.ProductServiceUsage=2]="ProductServiceUsage",e[e.ProductServicePerformance=4]="ProductServicePerformance",e[e.DeviceConfiguration=8]="DeviceConfiguration",e[e.InkingTypingSpeech=16]="InkingTypingSpeech"}(z||(z={})),function(e){e[e.ReservedDoNotUse=0]="ReservedDoNotUse",e[e.BasicEvent=10]="BasicEvent",e[e.FullEvent=100]="FullEvent",e[e.NecessaryServiceDataEvent=110]="NecessaryServiceDataEvent",e[e.AlwaysOnNecessaryServiceDataEvent=120]="AlwaysOnNecessaryServiceDataEvent"}(W||(W={})),function(e){e[e.Aria=0]="Aria",e[e.Nexus=1]="Nexus"}(Z||(Z={})),function(e){var t={},n={},r={};function i(e){if("object"!=typeof e)throw new Error("tokenTree must be an object");r=function e(t,n){if("object"!=typeof n)return n;for(var r=0,i=Object.keys(n);r<i.length;r++){var o=i[r];o in t&&(t[o],1)?t[o]=e(t[o],n[o]):t[o]=n[o]}return t}(r,e)}function o(e){if(t[e])return t[e];var n=s(e,Z.Aria);return"string"==typeof n?(t[e]=n,n):void 0}function a(e){if(n[e])return n[e];var t=s(e,Z.Nexus);return"number"==typeof t?(n[e]=t,t):void 0}function s(e,t){var n=e.split("."),i=r,o=void 0;if(i){for(var a=0;a<n.length-1;a++)i[n[a]]&&(i=i[n[a]],t===Z.Aria&&"string"==typeof i.ariaTenantToken?o=i.ariaTenantToken:t===Z.Nexus&&"number"==typeof i.nexusTenantToken&&(o=i.nexusTenantToken));return o}}e.setTenantToken=function(e,t,n){var r=e.split(".");if(r.length<2||"Office"!==r[0])V(O.Error,P.Core,(function(){return"Invalid namespace: "+e}));else{var o=Object.create(Object.prototype);t&&(o.ariaTenantToken=t),n&&(o.nexusTenantToken=n);var a,s=o;for(a=r.length-1;a>=0;--a){var u=Object.create(Object.prototype);u[r[a]]=s,s=u}i(s)}},e.setTenantTokens=i,e.getTenantTokens=function(e){var t=o(e),n=a(e);if(!n||!t)throw new Error("Could not find tenant token");return{ariaTenantToken:t,nexusTenantToken:n}},e.getAriaTenantToken=o,e.getNexusTenantToken=a,e.clear=function(){t={},n={},r={}}}(X||(X={}));var K,Q="3.1.24",Y=function(){function e(e,t){var n,r;this.onSendEvent=new D,this.persistentDataFields=[],e?(this.onSendEvent=e.onSendEvent,(n=this.persistentDataFields).push.apply(n,e.persistentDataFields)):this.persistentDataFields.push(S("OTelJS.Version",Q)),t&&(r=this.persistentDataFields).push.apply(r,t)}return e.prototype.sendTelemetryEvent=function(e){try{if(0===this.onSendEvent.getListenerCount())return void V(O.Warning,P.Core,(function(){return"No telemetry sinks are attached."}));var t=this.cloneEvent(e);this.processTelemetryEvent(t),this.onSendEvent.fireEvent(t)}catch(e){var n;n=e instanceof Error?e.message:JSON.stringify(e),V(O.Error,P.Core,(function(){return n}))}},e.prototype.processTelemetryEvent=function(e){var t;e.telemetryProperties||(e.telemetryProperties=X.getTenantTokens(e.eventName)),(t=e.dataFields).push.apply(t,this.persistentDataFields),i.validateTelemetryEvent(e)},e.prototype.addSink=function(e){this.onSendEvent.addListener((function(t){return e.sendTelemetryEvent(t)}))},e.prototype.setTenantToken=function(e,t,n){X.setTenantToken(e,t,n)},e.prototype.setTenantTokens=function(e){X.setTenantTokens(e)},e.prototype.cloneEvent=function(e){var t={eventName:e.eventName,eventFlags:e.eventFlags};return e.telemetryProperties&&(t.telemetryProperties={ariaTenantToken:e.telemetryProperties.ariaTenantToken,nexusTenantToken:e.telemetryProperties.nexusTenantToken}),e.eventContract&&(t.eventContract={name:e.eventContract.name,dataFields:e.eventContract.dataFields.slice()}),t.dataFields=e.dataFields?e.dataFields.slice():[],t},e}(),ee=(K=function(e,t){return(K=Object.setPrototypeOf||{__proto__:[]}instanceof Array&&function(e,t){e.__proto__=t}||function(e,t){for(var n in t)t.hasOwnProperty(n)&&(e[n]=t[n])})(e,t)},function(e,t){function n(){this.constructor=e}K(e,t),e.prototype=null===t?Object.create(t):(n.prototype=t.prototype,new n)}),te=function(e,t,n,r){return new(n||(n=Promise))((function(i,o){function a(e){try{u(r.next(e))}catch(e){o(e)}}function s(e){try{u(r.throw(e))}catch(e){o(e)}}function u(e){var t;e.done?i(e.value):(t=e.value,t instanceof n?t:new n((function(e){e(t)}))).then(a,s)}u((r=r.apply(e,t||[])).next())}))},ne=function(e,t){var n,r,i,o,a={label:0,sent:function(){if(1&i[0])throw i[1];return i[1]},trys:[],ops:[]};return o={next:s(0),throw:s(1),return:s(2)},"function"==typeof Symbol&&(o[Symbol.iterator]=function(){return this}),o;function s(o){return function(s){return function(o){if(n)throw new TypeError("Generator is already executing.");for(;a;)try{if(n=1,r&&(i=2&o[0]?r.return:o[0]?r.throw||((i=r.return)&&i.call(r),0):r.next)&&!(i=i.call(r,o[1])).done)return i;switch(r=0,i&&(o=[2&o[0],i.value]),o[0]){case 0:case 1:i=o;break;case 4:return a.label++,{value:o[1],done:!1};case 5:a.label++,r=o[1],o=[0];continue;case 7:o=a.ops.pop(),a.trys.pop();continue;default:if(!(i=(i=a.trys).length>0&&i[i.length-1])&&(6===o[0]||2===o[0])){a=0;continue}if(3===o[0]&&(!i||o[1]>i[0]&&o[1]<i[3])){a.label=o[1];break}if(6===o[0]&&a.label<i[1]){a.label=i[1],i=o;break}if(i&&a.label<i[2]){a.label=i[2],a.ops.push(o);break}i[2]&&a.ops.pop(),a.trys.pop();continue}o=t.call(e,a)}catch(e){o=[6,e],r=0}finally{n=i=0}if(5&o[0])throw o[1];return{value:o[0]?o[1]:void 0,done:!0}}([o,s])}}},re=function(e){function t(){return null!==e&&e.apply(this,arguments)||this}return ee(t,e),t.prototype.executeActivityAsync=function(e,t){return te(this,void 0,void 0,(function(){return ne(this,(function(n){return[2,this.createNewActivity(e).executeAsync(t)]}))}))},t.prototype.executeActivitySync=function(e,t){return this.createNewActivity(e).executeSync(t)},t.prototype.createNewActivity=function(e){return $.createNew(this,e)},t.prototype.sendActivity=function(e,t,n,r){return this.sendTelemetryEvent({eventName:e,eventContract:{name:b.Office.System.Activity.contractName,dataFields:b.Office.System.Activity.getFields(t)},dataFields:n,eventFlags:r})},t.prototype.sendError=function(e){var t=y.getFields("Error",e.error);return null!=e.dataFields&&t.push.apply(t,e.dataFields),this.sendTelemetryEvent({eventName:e.eventName,dataFields:t,eventFlags:e.eventFlags})},t}(Y);n.d(t,"Contracts",(function(){return b})),n.d(t,"ActivityScope",(function(){return $})),n.d(t,"getFieldsForContract",(function(){return T})),n.d(t,"addContractField",(function(){return F})),n.d(t,"DataClassification",(function(){return B})),n.d(t,"makeBooleanDataField",(function(){return h})),n.d(t,"makeInt64DataField",(function(){return g})),n.d(t,"makeDoubleDataField",(function(){return m})),n.d(t,"makeStringDataField",(function(){return S})),n.d(t,"DataFieldType",(function(){return r})),n.d(t,"getEffectiveEventFlags",(function(){return q})),n.d(t,"SamplingPolicy",(function(){return J})),n.d(t,"PersistencePriority",(function(){return R})),n.d(t,"CostPriority",(function(){return G})),n.d(t,"DataCategories",(function(){return z})),n.d(t,"DiagnosticLevel",(function(){return W})),n.d(t,"LogLevel",(function(){return O})),n.d(t,"Category",(function(){return P})),n.d(t,"onNotification",(function(){return j})),n.d(t,"logNotification",(function(){return V})),n.d(t,"SuppressNexus",(function(){return-1})),n.d(t,"SimpleTelemetryLogger",(function(){return Y})),n.d(t,"TelemetryLogger",(function(){return re}))}]);
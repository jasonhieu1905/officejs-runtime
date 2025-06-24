/******/ (function() { // webpackBootstrap
/*!**********************************!*\
  !*** ./src/commands/commands.ts ***!
  \**********************************/
function _typeof(o) { "@babel/helpers - typeof"; return _typeof = "function" == typeof Symbol && "symbol" == typeof Symbol.iterator ? function (o) { return typeof o; } : function (o) { return o && "function" == typeof Symbol && o.constructor === Symbol && o !== Symbol.prototype ? "symbol" : typeof o; }, _typeof(o); }
function _regeneratorRuntime() { "use strict"; /*! regenerator-runtime -- Copyright (c) 2014-present, Facebook, Inc. -- license (MIT): https://github.com/facebook/regenerator/blob/main/LICENSE */ _regeneratorRuntime = function _regeneratorRuntime() { return e; }; var t, e = {}, r = Object.prototype, n = r.hasOwnProperty, o = Object.defineProperty || function (t, e, r) { t[e] = r.value; }, i = "function" == typeof Symbol ? Symbol : {}, a = i.iterator || "@@iterator", c = i.asyncIterator || "@@asyncIterator", u = i.toStringTag || "@@toStringTag"; function define(t, e, r) { return Object.defineProperty(t, e, { value: r, enumerable: !0, configurable: !0, writable: !0 }), t[e]; } try { define({}, ""); } catch (t) { define = function define(t, e, r) { return t[e] = r; }; } function wrap(t, e, r, n) { var i = e && e.prototype instanceof Generator ? e : Generator, a = Object.create(i.prototype), c = new Context(n || []); return o(a, "_invoke", { value: makeInvokeMethod(t, r, c) }), a; } function tryCatch(t, e, r) { try { return { type: "normal", arg: t.call(e, r) }; } catch (t) { return { type: "throw", arg: t }; } } e.wrap = wrap; var h = "suspendedStart", l = "suspendedYield", f = "executing", s = "completed", y = {}; function Generator() {} function GeneratorFunction() {} function GeneratorFunctionPrototype() {} var p = {}; define(p, a, function () { return this; }); var d = Object.getPrototypeOf, v = d && d(d(values([]))); v && v !== r && n.call(v, a) && (p = v); var g = GeneratorFunctionPrototype.prototype = Generator.prototype = Object.create(p); function defineIteratorMethods(t) { ["next", "throw", "return"].forEach(function (e) { define(t, e, function (t) { return this._invoke(e, t); }); }); } function AsyncIterator(t, e) { function invoke(r, o, i, a) { var c = tryCatch(t[r], t, o); if ("throw" !== c.type) { var u = c.arg, h = u.value; return h && "object" == _typeof(h) && n.call(h, "__await") ? e.resolve(h.__await).then(function (t) { invoke("next", t, i, a); }, function (t) { invoke("throw", t, i, a); }) : e.resolve(h).then(function (t) { u.value = t, i(u); }, function (t) { return invoke("throw", t, i, a); }); } a(c.arg); } var r; o(this, "_invoke", { value: function value(t, n) { function callInvokeWithMethodAndArg() { return new e(function (e, r) { invoke(t, n, e, r); }); } return r = r ? r.then(callInvokeWithMethodAndArg, callInvokeWithMethodAndArg) : callInvokeWithMethodAndArg(); } }); } function makeInvokeMethod(e, r, n) { var o = h; return function (i, a) { if (o === f) throw Error("Generator is already running"); if (o === s) { if ("throw" === i) throw a; return { value: t, done: !0 }; } for (n.method = i, n.arg = a;;) { var c = n.delegate; if (c) { var u = maybeInvokeDelegate(c, n); if (u) { if (u === y) continue; return u; } } if ("next" === n.method) n.sent = n._sent = n.arg;else if ("throw" === n.method) { if (o === h) throw o = s, n.arg; n.dispatchException(n.arg); } else "return" === n.method && n.abrupt("return", n.arg); o = f; var p = tryCatch(e, r, n); if ("normal" === p.type) { if (o = n.done ? s : l, p.arg === y) continue; return { value: p.arg, done: n.done }; } "throw" === p.type && (o = s, n.method = "throw", n.arg = p.arg); } }; } function maybeInvokeDelegate(e, r) { var n = r.method, o = e.iterator[n]; if (o === t) return r.delegate = null, "throw" === n && e.iterator.return && (r.method = "return", r.arg = t, maybeInvokeDelegate(e, r), "throw" === r.method) || "return" !== n && (r.method = "throw", r.arg = new TypeError("The iterator does not provide a '" + n + "' method")), y; var i = tryCatch(o, e.iterator, r.arg); if ("throw" === i.type) return r.method = "throw", r.arg = i.arg, r.delegate = null, y; var a = i.arg; return a ? a.done ? (r[e.resultName] = a.value, r.next = e.nextLoc, "return" !== r.method && (r.method = "next", r.arg = t), r.delegate = null, y) : a : (r.method = "throw", r.arg = new TypeError("iterator result is not an object"), r.delegate = null, y); } function pushTryEntry(t) { var e = { tryLoc: t[0] }; 1 in t && (e.catchLoc = t[1]), 2 in t && (e.finallyLoc = t[2], e.afterLoc = t[3]), this.tryEntries.push(e); } function resetTryEntry(t) { var e = t.completion || {}; e.type = "normal", delete e.arg, t.completion = e; } function Context(t) { this.tryEntries = [{ tryLoc: "root" }], t.forEach(pushTryEntry, this), this.reset(!0); } function values(e) { if (e || "" === e) { var r = e[a]; if (r) return r.call(e); if ("function" == typeof e.next) return e; if (!isNaN(e.length)) { var o = -1, i = function next() { for (; ++o < e.length;) if (n.call(e, o)) return next.value = e[o], next.done = !1, next; return next.value = t, next.done = !0, next; }; return i.next = i; } } throw new TypeError(_typeof(e) + " is not iterable"); } return GeneratorFunction.prototype = GeneratorFunctionPrototype, o(g, "constructor", { value: GeneratorFunctionPrototype, configurable: !0 }), o(GeneratorFunctionPrototype, "constructor", { value: GeneratorFunction, configurable: !0 }), GeneratorFunction.displayName = define(GeneratorFunctionPrototype, u, "GeneratorFunction"), e.isGeneratorFunction = function (t) { var e = "function" == typeof t && t.constructor; return !!e && (e === GeneratorFunction || "GeneratorFunction" === (e.displayName || e.name)); }, e.mark = function (t) { return Object.setPrototypeOf ? Object.setPrototypeOf(t, GeneratorFunctionPrototype) : (t.__proto__ = GeneratorFunctionPrototype, define(t, u, "GeneratorFunction")), t.prototype = Object.create(g), t; }, e.awrap = function (t) { return { __await: t }; }, defineIteratorMethods(AsyncIterator.prototype), define(AsyncIterator.prototype, c, function () { return this; }), e.AsyncIterator = AsyncIterator, e.async = function (t, r, n, o, i) { void 0 === i && (i = Promise); var a = new AsyncIterator(wrap(t, r, n, o), i); return e.isGeneratorFunction(r) ? a : a.next().then(function (t) { return t.done ? t.value : a.next(); }); }, defineIteratorMethods(g), define(g, u, "Generator"), define(g, a, function () { return this; }), define(g, "toString", function () { return "[object Generator]"; }), e.keys = function (t) { var e = Object(t), r = []; for (var n in e) r.push(n); return r.reverse(), function next() { for (; r.length;) { var t = r.pop(); if (t in e) return next.value = t, next.done = !1, next; } return next.done = !0, next; }; }, e.values = values, Context.prototype = { constructor: Context, reset: function reset(e) { if (this.prev = 0, this.next = 0, this.sent = this._sent = t, this.done = !1, this.delegate = null, this.method = "next", this.arg = t, this.tryEntries.forEach(resetTryEntry), !e) for (var r in this) "t" === r.charAt(0) && n.call(this, r) && !isNaN(+r.slice(1)) && (this[r] = t); }, stop: function stop() { this.done = !0; var t = this.tryEntries[0].completion; if ("throw" === t.type) throw t.arg; return this.rval; }, dispatchException: function dispatchException(e) { if (this.done) throw e; var r = this; function handle(n, o) { return a.type = "throw", a.arg = e, r.next = n, o && (r.method = "next", r.arg = t), !!o; } for (var o = this.tryEntries.length - 1; o >= 0; --o) { var i = this.tryEntries[o], a = i.completion; if ("root" === i.tryLoc) return handle("end"); if (i.tryLoc <= this.prev) { var c = n.call(i, "catchLoc"), u = n.call(i, "finallyLoc"); if (c && u) { if (this.prev < i.catchLoc) return handle(i.catchLoc, !0); if (this.prev < i.finallyLoc) return handle(i.finallyLoc); } else if (c) { if (this.prev < i.catchLoc) return handle(i.catchLoc, !0); } else { if (!u) throw Error("try statement without catch or finally"); if (this.prev < i.finallyLoc) return handle(i.finallyLoc); } } } }, abrupt: function abrupt(t, e) { for (var r = this.tryEntries.length - 1; r >= 0; --r) { var o = this.tryEntries[r]; if (o.tryLoc <= this.prev && n.call(o, "finallyLoc") && this.prev < o.finallyLoc) { var i = o; break; } } i && ("break" === t || "continue" === t) && i.tryLoc <= e && e <= i.finallyLoc && (i = null); var a = i ? i.completion : {}; return a.type = t, a.arg = e, i ? (this.method = "next", this.next = i.finallyLoc, y) : this.complete(a); }, complete: function complete(t, e) { if ("throw" === t.type) throw t.arg; return "break" === t.type || "continue" === t.type ? this.next = t.arg : "return" === t.type ? (this.rval = this.arg = t.arg, this.method = "return", this.next = "end") : "normal" === t.type && e && (this.next = e), y; }, finish: function finish(t) { for (var e = this.tryEntries.length - 1; e >= 0; --e) { var r = this.tryEntries[e]; if (r.finallyLoc === t) return this.complete(r.completion, r.afterLoc), resetTryEntry(r), y; } }, catch: function _catch(t) { for (var e = this.tryEntries.length - 1; e >= 0; --e) { var r = this.tryEntries[e]; if (r.tryLoc === t) { var n = r.completion; if ("throw" === n.type) { var o = n.arg; resetTryEntry(r); } return o; } } throw Error("illegal catch attempt"); }, delegateYield: function delegateYield(e, r, n) { return this.delegate = { iterator: values(e), resultName: r, nextLoc: n }, "next" === this.method && (this.arg = t), y; } }, e; }
function asyncGeneratorStep(n, t, e, r, o, a, c) { try { var i = n[a](c), u = i.value; } catch (n) { return void e(n); } i.done ? t(u) : Promise.resolve(u).then(r, o); }
function _asyncToGenerator(n) { return function () { var t = this, e = arguments; return new Promise(function (r, o) { var a = n.apply(t, e); function _next(n) { asyncGeneratorStep(a, r, o, _next, _throw, "next", n); } function _throw(n) { asyncGeneratorStep(a, r, o, _next, _throw, "throw", n); } _next(void 0); }); }; }
/*
 * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */

// import { loginSilentAndGetAccessToken } from "./get-user-data";

/* global global, Office, self, window */

Office.onReady(function () {
  // If needed, Office.js is ready to be called
});
function changeHeader(_x) {
  return _changeHeader.apply(this, arguments);
}
function _changeHeader() {
  _changeHeader = _asyncToGenerator(/*#__PURE__*/_regeneratorRuntime().mark(function _callee2(event) {
    return _regeneratorRuntime().wrap(function _callee2$(_context2) {
      while (1) switch (_context2.prev = _context2.next) {
        case 0:
          _context2.next = 2;
          return Word.run(/*#__PURE__*/function () {
            var _ref = _asyncToGenerator(/*#__PURE__*/_regeneratorRuntime().mark(function _callee(context) {
              var body, xmlText, xmlPara, placeholder, cc, response, profile, msg;
              return _regeneratorRuntime().wrap(function _callee$(_context) {
                while (1) switch (_context.prev = _context.next) {
                  case 0:
                    // 1) Grab the document body
                    body = context.document.body; // 2) Fetch your custom XML (or capture the error)
                    _context.prev = 1;
                    _context.next = 4;
                    return getCustomXmlPart();
                  case 4:
                    xmlText = _context.sent;
                    _context.next = 10;
                    break;
                  case 7:
                    _context.prev = 7;
                    _context.t0 = _context["catch"](1);
                    xmlText = "Error getting custom XML part: " + JSON.stringify(_context.t0);
                  case 10:
                    // 3) Insert the XML (or error) at the end, color it green
                    xmlPara = body.insertParagraph(xmlText, Word.InsertLocation.end);
                    xmlPara.font.color = "#07641d";

                    // 4) Blank line for separation
                    body.insertParagraph("", Word.InsertLocation.end);

                    // 5) Create a placeholder paragraph and wrap it in a content control
                    placeholder = body.insertParagraph("", Word.InsertLocation.end);
                    cc = placeholder.insertContentControl();
                    cc.tag = "testHtmlCc";
                    cc.title = "Test HTML Content Control";
                    cc.appearance = Word.ContentControlAppearance.boundingBox;

                    // 6) Insert your HTML into *that* content control, replacing the placeholder
                    cc.insertHtml("<h3 style='color: red'>Hello doan 2</h3>", Word.InsertLocation.replace);

                    // 8) Sync once at the end
                    _context.next = 21;
                    return context.sync();
                  case 21:
                    _context.prev = 21;
                    if (!(typeof fetch !== "function")) {
                      _context.next = 25;
                      break;
                    }
                    body.insertParagraph("Fetch API is not available in this environment.", Word.InsertLocation.end);
                    return _context.abrupt("return");
                  case 25:
                    _context.next = 27;
                    return fetch("https://graph.microsoft.com/v1.0/me", {
                      headers: {
                        Authorization: "Bearer eyJ0eXAiOiJKV1QiLCJub25jZSI6IkZMMC1Oc1BTUnB0RGVaOERlZHc4VlZ4V1k3RnhNX1B1ay1WbTc2Z0hyOEEiLCJhbGciOiJSUzI1NiIsIng1dCI6IkNOdjBPSTNSd3FsSEZFVm5hb01Bc2hDSDJYRSIsImtpZCI6IkNOdjBPSTNSd3FsSEZFVm5hb01Bc2hDSDJYRSJ9.eyJhdWQiOiIwMDAwMDAwMy0wMDAwLTAwMDAtYzAwMC0wMDAwMDAwMDAwMDAiLCJpc3MiOiJodHRwczovL3N0cy53aW5kb3dzLm5ldC9mM2Q2Zjc1Zi02YjQwLTQ1MmYtYWU5YS02MjI2M2YyNmM2MzkvIiwiaWF0IjoxNzUwNzUwNzM3LCJuYmYiOjE3NTA3NTA3MzcsImV4cCI6MTc1MDc1NjM4MCwiYWNjdCI6MCwiYWNyIjoiMSIsImFjcnMiOlsicDEiXSwiYWlvIjoiQVdRQW0vOFpBQUFBTFUrbklDK3pLdTFjeDZkZDhTWEl1OGg0VEFwanUvM2xjOUkvVkF4eEJoY0tJZmVkZmZIUEsyTlV4NFgzQnZJL2NCbHMvODIvRlRjNE9ST1ZVZFZSNWlJZnNONzZNZ3lOOEJzTzJVVFRCTjhYWHpyWWhZUDBBQ2NKMnZTU2pTcC8iLCJhbXIiOlsicHdkIiwibWZhIl0sImFwcF9kaXNwbGF5bmFtZSI6Ik9mZmljZS1BZGQtaW4tU1NPLU5BQSIsImFwcGlkIjoiZTE1ZGY1NWItNzkzNC00ZThhLWIyZWYtYmRjMWRjMjIwNGI1IiwiYXBwaWRhY3IiOiIwIiwiZmFtaWx5X25hbWUiOiJEb2FuIiwiZ2l2ZW5fbmFtZSI6IkhpZXUiLCJpZHR5cCI6InVzZXIiLCJpcGFkZHIiOiI0Mi4xMTUuMTkzLjQzIiwibmFtZSI6IkhpZXUgRG9hbiIsIm9pZCI6IjIyNDdiMjliLTI1YzgtNDFkMy1hMWJjLWYzMDhmOTVkYmUwMCIsInBsYXRmIjoiMyIsInB1aWQiOiIxMDAzMjAwMTVGNkQzNzBGIiwicHdkX3VybCI6Imh0dHBzOi8vcG9ydGFsLm1pY3Jvc29mdG9ubGluZS5jb20vQ2hhbmdlUGFzc3dvcmQuYXNweCIsInJoIjoiMS5BWEVBWF9mVzgwQnJMMFd1bW1JbVB5YkdPUU1BQUFBQUFBQUF3QUFBQUFBQUFBQnhBRWR4QUEuIiwic2NwIjoiRmlsZXMuUmVhZCBvcGVuaWQgcHJvZmlsZSBVc2VyLlJlYWQgZW1haWwiLCJzaWQiOiIwMDVmMTdkOS0zNTY3LWZhYzEtYmZmOS01ZDgwMjFkNGY4MGMiLCJzdWIiOiJJeHNBb09rWEJTVlpMY1NGamo5YkVkcnp5UTdXekd1dnVZVXM4QzNKRWVRIiwidGVuYW50X3JlZ2lvbl9zY29wZSI6IkFTIiwidGlkIjoiZjNkNmY3NWYtNmI0MC00NTJmLWFlOWEtNjIyNjNmMjZjNjM5IiwidW5pcXVlX25hbWUiOiJqYXNvbi5oaWV1QGhpZXVkb2FuZGV2Lm9ubWljcm9zb2Z0LmNvbSIsInVwbiI6Imphc29uLmhpZXVAaGlldWRvYW5kZXYub25taWNyb3NvZnQuY29tIiwidXRpIjoiaFVNYXJGSm9VVUNXLWJjLWpqS2tBQSIsInZlciI6IjEuMCIsIndpZHMiOlsiNjJlOTAzOTQtNjlmNS00MjM3LTkxOTAtMDEyMTc3MTQ1ZTEwIiwiYjc5ZmJmNGQtM2VmOS00Njg5LTgxNDMtNzZiMTk0ZTg1NTA5Il0sInhtc19mdGQiOiIxajJQRHNoVnNtNVd4eG1xMVNlcmg3ZkkxUHRyZ3NoWlNWamNXMVJydFNNQmFtRndZVzVsWVhOMExXUnpiWE0iLCJ4bXNfaWRyZWwiOiIxNiAxIiwieG1zX3N0Ijp7InN1YiI6Ikh0bFkzbHVqS1gtYjVMRWVkN0VPX3Z0T2YwbWVGZmNZYlgzaTNKeUpPaWsifSwieG1zX3RjZHQiOjE2MjY0MjMyNDR9.XwqHtf1Ku8-ViJ9BApJ-OecpEy6PlgC3vW5M0Ur73mJVulgeHfLCtgyos5ANFSVRMrWDyC3cgzOgfuGJisZTA0EvNK1m304d3orl6HqgeMRYqzszHP4STDYAVG2m1dKGzXjYM0uFrsa-ibWhvt5_xrai__lLzPDgXqCBxrHTbaXxIC6Nb6Q5SuO9qxaNjKFxuGWoCIXPpEt7Pf69fTBdK0KV6YUDPFwgTqcCCx-kMzm04TjahWkUn3yi5cF5-RqOxBTnHYHU6GgOS_Df7CR_PcMh6ecwUHIYcgw0SFx7kcM_bj-7-lML1QAYDgxQwPRvNPrwlilQPKNEMtQl3HWtWQ",
                        Accept: "application/json"
                      }
                    });
                  case 27:
                    response = _context.sent;
                    if (response.ok) {
                      _context.next = 32;
                      break;
                    }
                    body.insertParagraph("Failed to fetch user profile: ".concat(response.statusText), Word.InsertLocation.end);
                    _context.next = 36;
                    break;
                  case 32:
                    _context.next = 34;
                    return response.json();
                  case 34:
                    profile = _context.sent;
                    body.insertParagraph("User: ".concat(profile.displayName, " (").concat(profile.mail || profile.userPrincipalName, ")"), Word.InsertLocation.end);
                  case 36:
                    _context.next = 42;
                    break;
                  case 38:
                    _context.prev = 38;
                    _context.t1 = _context["catch"](21);
                    msg = _context.t1 && _context.t1.message ? _context.t1.message : "Unknown error fetching user profile.";
                    body.insertParagraph("Error fetching user profile : ".concat(msg), Word.InsertLocation.end);
                  case 42:
                  case "end":
                    return _context.stop();
                }
              }, _callee, null, [[1, 7], [21, 38]]);
            }));
            return function (_x2) {
              return _ref.apply(this, arguments);
            };
          }());
        case 2:
          // Calling event.completed is required. event.completed lets the platform know that processing has completed.
          event.completed();
        case 3:
        case "end":
          return _context2.stop();
      }
    }, _callee2);
  }));
  return _changeHeader.apply(this, arguments);
}
function getCustomXmlPart() {
  return _getCustomXmlPart.apply(this, arguments);
} // The add-in command functions need to be available in global scope
function _getCustomXmlPart() {
  _getCustomXmlPart = _asyncToGenerator(/*#__PURE__*/_regeneratorRuntime().mark(function _callee3() {
    return _regeneratorRuntime().wrap(function _callee3$(_context3) {
      while (1) switch (_context3.prev = _context3.next) {
        case 0:
          return _context3.abrupt("return", new Promise(function (resolve, reject) {
            Office.context.document.customXmlParts.getByNamespaceAsync("http://schemas.openxmlformats.org/package/2006/metadata/core-properties", {}, function (result) {
              if (result.status === Office.AsyncResultStatus.Failed) {
                reject(result.error);
                return;
              }
              var customXmlPart = result.value[0];
              if (!customXmlPart) {
                resolve(undefined);
                return;
              }
              customXmlPart.getXmlAsync({}, function (xmlResult) {
                if (xmlResult.status === Office.AsyncResultStatus.Failed) {
                  reject(xmlResult.error);
                  return;
                }
                resolve(xmlResult.value);
              });
            });
          }));
        case 1:
        case "end":
          return _context3.stop();
      }
    }, _callee3);
  }));
  return _getCustomXmlPart.apply(this, arguments);
}
Office.actions.associate("changeHeader", changeHeader);
/******/ })()
;
//# sourceMappingURL=commands.js.map

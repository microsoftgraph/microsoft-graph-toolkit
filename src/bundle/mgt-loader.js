(function() {
  "use strict";

  var loaded = false;
  var pwd = getpath();

  // load promise polyfill if needed
  if (!window.Promise) {
    addScript("https://cdn.jsdelivr.net/npm/es6-promise@4/dist/es6-promise.js");
    addScript(
      "https://cdn.jsdelivr.net/npm/es6-promise@4/dist/es6-promise.auto.js"
    );
  }

  polyfillFind();

  window.WebComponents = window.WebComponents || {};
  window.WebComponents.root = pwd + "wc/";

  //load webcomponent polyfill
  addScript(pwd + "wc/webcomponents-loader.js", function() {
    WebComponents.waitFor(function() {
      if (loaded) {
        return;
      }
      loaded = true;

      // decide es5 or es6 --TODO
      if (es6()) {
        addScript(pwd + "mgt.es6.js");
      } else {
        addScript(pwd + "mgt.es5.js");
      }
    });
  });

  function getpath() {
    var scripts = document.getElementsByTagName("script");
    var path = scripts[scripts.length - 1].src.split("?")[0];
    var dir =
      path
        .split("/")
        .slice(0, -1)
        .join("/") + "/";
    return dir;
  }

  // from https://stackoverflow.com/questions/29046635/javascript-es6-cross-browser-detection
  function es6() {
    "use strict";

    if (typeof Symbol == "undefined") return false;
    try {
      eval("class Foo {}");
      eval("var bar = (x) => x+1");
    } catch (e) {
      return false;
    }

    return true;
  }

  function addScript(src, onload) {
    var tag = document.createElement("script");

    if (onload) {
      tag.addEventListener("load", onload);
    }

    tag.src = src;
    document.head.appendChild(tag);
  }

  function polyfillFind() {
    // https://tc39.github.io/ecma262/#sec-array.prototype.find
    if (!Array.prototype.find) {
      Object.defineProperty(Array.prototype, "find", {
        value: function(predicate) {
          // 1. Let O be ? ToObject(this value).
          if (this == null) {
            throw new TypeError('"this" is null or not defined');
          }

          var o = Object(this);

          // 2. Let len be ? ToLength(? Get(O, "length")).
          var len = o.length >>> 0;

          // 3. If IsCallable(predicate) is false, throw a TypeError exception.
          if (typeof predicate !== "function") {
            throw new TypeError("predicate must be a function");
          }

          // 4. If thisArg was supplied, let T be thisArg; else let T be undefined.
          var thisArg = arguments[1];

          // 5. Let k be 0.
          var k = 0;

          // 6. Repeat, while k < len
          while (k < len) {
            // a. Let Pk be ! ToString(k).
            // b. Let kValue be ? Get(O, Pk).
            // c. Let testResult be ToBoolean(? Call(predicate, T, « kValue, k, O »)).
            // d. If testResult is true, return kValue.
            var kValue = o[k];
            if (predicate.call(thisArg, kValue, k, o)) {
              return kValue;
            }
            // e. Increase k by 1.
            k++;
          }

          // 7. Return undefined.
          return undefined;
        },
        configurable: true,
        writable: true
      });
    }
  }
})();

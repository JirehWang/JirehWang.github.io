/**
 * api.js - GAS API 橋接層
 */
if (typeof window.GAS_CONFIG === 'undefined') {
  console.error('❌ 找不到 config.js');
}

window.google = {
  script: {
    url: {
      getLocation: function(callback) {
        const params = {};
        new URLSearchParams(window.location.search).forEach((v, k) => { params[k] = v; });
        callback({ parameter: params });
      }
    },
    run: (function() {
      // 建立一個帶有 successHandler / failureHandler 的呼叫鏈
      function makeRunner(successHandler, failureHandler) {
        const handler = {
          withSuccessHandler: function(fn) {
            return makeRunner(fn, failureHandler);
          },
          withFailureHandler: function(fn) {
            return makeRunner(successHandler, fn);
          }
        };

        // 用 Proxy 攔截所有函式名稱
        return new Proxy(handler, {
          get(target, functionName) {
            // 如果是 withSuccessHandler / withFailureHandler 就直接回傳
            if (functionName in target) return target[functionName];

            // 否則視為 GAS 函式名稱，回傳一個可以呼叫的函式
            return function(...args) {
              const apiUrl = window.GAS_CONFIG && window.GAS_CONFIG.apiUrl;
              if (!apiUrl || apiUrl.includes('YOUR_SCRIPT_ID')) {
                const err = new Error('❌ 請先設定 config.js 中的 GAS Web App URL！');
                console.error(err.message);
                if (failureHandler) failureHandler(err);
                return;
              }

              fetch(apiUrl, {
                method: 'POST',
                headers: { 'Content-Type': 'text/plain' },
                body: JSON.stringify({
                  action: functionName,
                  payload: args.length === 1 ? args[0] : args
                })
              })
              .then(res => res.json())
              .then(data => {
                if (data.error) {
                  if (failureHandler) failureHandler(new Error(data.error));
                } else {
                  if (successHandler) successHandler(data.data);
                }
              })
              .catch(err => {
                if (failureHandler) failureHandler(err);
              });
            };
          }
        });
      }

      // run 本身就是一個 makeRunner 的起點
      return makeRunner(null, (err) => console.error('GAS Error:', err));
    })()
  }
};

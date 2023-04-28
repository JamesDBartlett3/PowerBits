// ==UserScript==
// @name         Custom Power BI Page Titles
// @namespace    http://tampermonkey.net/
// @version      0.1
// @description  Specify custom titles for specific pages in Power BI, so the browser tab will actually say what's on that page instead of just "Power BI."
// @author       @JamesDBartlett3 (techhub.social/@JamesDBartlett3, github.com/JamesDBartlett3)
// @match        https://app.powerbi.com/groups/*/list
// @match        https://app.powerbi.com/groups/*/settings
// @match        https://app.powerbi.com/pipelines/*
// @icon         https://app.powerbi.com/images/PowerBI_Favicon.ico
// @grant        none
// ==/UserScript==

/*
  Installation & Usage Guide
  1) Install the Tampermonkey browser extension: https://www.tampermonkey.net
  2) Open this file in "Raw" mode on GitHub: https://github.com/JamesDBartlett3/PowerBits/raw/main/UserScripts/CustomPowerBIPageTitles.user.js
  3) Click the "Install" button
  4) Open the TamperMonkey dashboard and open the "Custom Power BI Page Titles" script for editing
  5) Edit the "tabNames" object below, adding the tab names you want on the left and the corresponding artifact IDs on the right
  6) Save your changes (Ctrl+S) and close the editor
  7) Refresh any Power BI tabs you have open
*/

(function () {
  "use strict";

  let tabNames = {
    "Exec[Pipe]": "ac8ae510-54f4-4046-a64c-58c7b2541b89",
    "Exec[Dev]": "ab86bc72-796d-41dc-9201-15954a866758",
    "Exec[Test]": "82f58e8b-cfd3-42f1-b723-ce199188d2fc",
    "Exec[Prod]": "0691620f-dab9-480b-816c-31c72880b903",
    "Admi[Pipe]": "073c8496-1820-437d-8071-ae277ecd8aa3",
    "Admi[Dev]": "b43f46d5-5c97-4a7a-9c7f-1ad0ef189365",
    "Admi[Test]": "8ba5087a-d778-4b10-8b50-36ce65304ef0",
    "Admi[Prod]": "d083e445-e431-47d8-970f-9f8723032404",
    _Dataflows: "f48aaaa7-3a82-4897-92b4-ef797db67f5f",
    _Datasets: "288311c7-879a-495c-b485-56b2fb347f50",
    _DMU_Store: "275d531f-04b0-4081-be1e-51537ec18b4f",
    _DMU_D2L: "c5208007-3114-420b-bb97-3e75ee613b86",
  };

  function getKeyByValue(object, value) {
    return Object.keys(object).find((key) => value.includes(object[key]));
  }

  function setTitle(t) {
    document.title = t;
  }

  for (var key in tabNames) {
    if (document.URL.includes(tabNames[key])) {
      setTitle(key);
    }
  }

  (function (history) {
    var pushState = history.pushState;
    history.pushState = function (state, key, path) {
      if (typeof history.onpushstate == "function") {
        history.onpushstate({
          state: state,
          path: path,
        });
      }
      pushState.apply(history, arguments);
    };
    window.onpopstate = history.onpushstate = function (e) {
      var pageName = getKeyByValue(tabNames, e.path);
      if (pageName) {
        document.title = pageName;
      }
    };
  })(window.history);
})();

// JS Fetch from loading the Power BI homepage
// TODO: Use this to fetch the current artifact name and set the title dynamically
/* fetch("https://wabi-us-north-central-h-primary-redirect.analysis.windows.net/metadata/app?preferReadOnlySession=true&cmd=home", {
  "headers": {
    "accept": "application/json",
    "accept-language": "en-US,en;q=0.9",
    "activityid": "xxx",
    "authorization": "Bearer XXXXX",
    "cache-control": "no-cache",
    "pragma": "no-cache",
    "requestid": "xxx",
    "sec-ch-ua": "\"Chromium\";v=\"112\", \"Microsoft Edge\";v=\"112\", \"Not:A-Brand\";v=\"99\"",
    "sec-ch-ua-mobile": "?0",
    "sec-ch-ua-platform": "\"Windows\"",
    "sec-fetch-dest": "empty",
    "sec-fetch-mode": "cors",
    "sec-fetch-site": "cross-site"
  },
  "referrer": "https://app.powerbi.com/",
  "referrerPolicy": "strict-origin-when-cross-origin",
  "body": null,
  "method": "GET",
  "mode": "cors",
  "credentials": "include"
}); */

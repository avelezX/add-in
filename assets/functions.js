/* global CustomFunctions */
console.log("Custom Functions script loaded");


var RPC_POST_URL = "https://tvpehjbqxpiswkqszwwv.supabase.co/rest/v1/rpc/search";
var PARAM_NAME   = "ticket";
var PROFILE_NAME = "public";
var API_KEY      = "eyJhbGciOiJIUzI1NiIsInR5cCI6IkpXVCJ9.eyJpc3MiOiJzdXBhYmFzZSIsInJlZiI6InR2cGVoamJxeHBpc3drcXN6d3d2Iiwicm9sZSI6ImFub24iLCJpYXQiOjE2OTY0NTEzODksImV4cCI6MjAxMjAyNzM4OX0.LZW0i9HU81lCdyjAdqjwwF4hkuSVtsJsSDQh7blzozw";

/** "YYYY-MM-DD" -> Excel serial (UTC) */
function ymdToExcelSerial(isoYmd) {
  var parts = String(isoYmd).split("-");
  var y = Number(parts[0]), m = Number(parts[1]) || 1, d = Number(parts[2]) || 1;
  var excelEpochUTC = Date.UTC(1899, 11, 30);
  var thisUTC = Date.UTC(y, m - 1, d);
  return (thisUTC - excelEpochUTC) / 86400000;
}

/** =XERENITY.XTY(ticker) -> [["time","value"], [serial,value], ...] */
function XTY(ticker) {
  return new Promise(function(resolve, reject) {
    if (!ticker || typeof ticker !== "string") {
      return reject(new CustomFunctions.Error(
        CustomFunctions.ErrorCode.invalidValue,
        'Debes pasar un ticker, ej: XERENITY.XTY("ibr_1yr").'
      ));
    }

    var headers = {
      "Content-Type": "application/json",
      "content-profile": PROFILE_NAME,
      "Accept-Profile": PROFILE_NAME,
      "apikey": API_KEY,
      "Authorization": "Bearer " + API_KEY
    };

    var bodyObj = {}; bodyObj[PARAM_NAME] = ticker.trim();

    fetch(RPC_POST_URL, {
      method: "POST",
      headers: headers,
      body: JSON.stringify(bodyObj)
    })
    .then(function(res) {
      if (!res.ok) {
        return res.text().then(function(t){ throw new Error("HTTP " + res.status + (t ? " - " + t : "")); });
      }
      return res.json();
    })
    .then(function(payload) {
      var rows;
      if (Array.isArray(payload)) {
        rows = payload;
      } else if (payload && Array.isArray(payload.data)) {
        rows = payload.data;
      } else {
        rows = [];
      }

      if (!rows.length) return resolve([["time","value"], [null, null]]);

      var out = [["time","value"]];
      for (var i = 0; i < rows.length; i++) {
        var r = rows[i];
        if (r && r.time != null && r.value != null) {
          out.push([ ymdToExcelSerial(r.time), Number(r.value) ]);
        }
      }
      if (out.length === 1) out.push([null, null]);
      resolve(out);
    })
    .catch(function(err) {
      reject(new CustomFunctions.Error(
        CustomFunctions.ErrorCode.connection,
        String(err && err.message ? err.message : err)
      ));
    });
  });
}

CustomFunctions.associate("XTY", XTY);

function PING() { return 1; }
CustomFunctions.associate("PING", PING);

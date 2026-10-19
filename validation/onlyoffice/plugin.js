/* global Api */
"use strict";

// Serialized into the editor context: keep this function independent of the UI.
function inspectPresentation() {
  function read(object, method) {
    if (!object || typeof object[method] !== "function") {
      return { status: "unavailable", method: method };
    }
    try {
      var value = object[method]();
      return value === undefined
        ? { status: "undefined", method: method }
        : { status: "ok", value: value };
    } catch (error) {
      return { status: "error", method: method, message: String(error) };
    }
  }
  function shapeRecord(shape) {
    var record = {
      id: read(shape, "GetInternalId"), kind: read(shape, "GetClassType"),
      name: read(shape, "GetName"),
      x: read(shape, "GetPosX"), y: read(shape, "GetPosY"),
      width: read(shape, "GetWidth"), height: read(shape, "GetHeight")
    };
    if (record.kind.value === "chart") {
      var series = read(shape, "GetSeries");
      record.seriesCount = series.status === "ok" && Array.isArray(series.value)
        ? { status: "ok", value: series.value.length }
        : { status: series.status === "ok" ? "unexpected-return-type" : series.status,
            message: series.message };
      record.hasSetSeriaValues = typeof shape.SetSeriaValues === "function";
      record.writeTest = "not-run";
    }
    if (record.kind.value === "group" || record.kind.value === "smartArt") {
      record.childrenInspection = "not-implemented; do not infer absence of children";
    }
    return record;
  }
  function shapeList(result) {
    if (result.status !== "ok") return result;
    if (!Array.isArray(result.value)) return { status: "unexpected-return-type" };
    return { status: "ok", value: result.value.map(shapeRecord) };
  }
  try {
    var presentation = Api.GetPresentation();
    var slide = read(presentation, "GetCurrentSlide");
    var selection = read(Api, "GetSelection");
    return JSON.stringify({
      schema: "intent-canvas-onlyoffice-probe/1",
      scope: "Read-only public API sample. A method's presence is not a successful edit test.",
      units: "EMU; group/screen coordinate mapping not verified",
      slideDrawings: slide.status === "ok"
        ? shapeList(read(slide.value, "GetAllDrawings")) : slide,
      selectionType: selection.status === "ok" ? read(selection.value, "GetType") : selection,
      selectedDrawings: selection.status === "ok"
        ? shapeList(read(selection.value, "GetShapes")) : selection
    });
  } catch (error) {
    return JSON.stringify({ status: "error", message: String(error) });
  }
}

(function () {
  if (!window.Asc || !window.Asc.plugin) return;
  var plugin = window.Asc.plugin;
  var snapshot = null;
  plugin.init = function () {
    var button = document.getElementById("inspect");
    button.disabled = false;
    button.onclick = function () {
      snapshot = null;
      document.getElementById("download").disabled = true;
      button.disabled = true;
      document.getElementById("result").textContent = "读取中…";
      var completed = false;
      var timeout = setTimeout(function () {
        completed = true;
        button.disabled = false;
        document.getElementById("result").textContent = "接口未返回结果；本次读取未完成，请检查宿主版本或重试。";
      }, 15000);
      function done(raw) {
        if (completed) return;
        completed = true;
        clearTimeout(timeout);
        button.disabled = false;
        try {
          if (typeof raw !== "string") throw new Error("No JSON result returned by host");
          snapshot = JSON.parse(raw);
          snapshot.capturedAt = new Date().toISOString();
          document.getElementById("result").textContent = JSON.stringify(snapshot, null, 2);
          document.getElementById("download").disabled = false;
        } catch (error) {
          document.getElementById("result").textContent = String(error);
        }
      }
      try {
        // Read-only command: no recalculation request, no edits, window remains open.
        plugin.callCommand(inspectPresentation, false, false, done);
      } catch (error) { done(JSON.stringify({ status: "error", message: String(error) })); }
    };
    document.getElementById("download").onclick = function () {
      if (!snapshot) return;
      var url = URL.createObjectURL(new Blob([JSON.stringify(snapshot, null, 2)], { type: "application/json" }));
      var link = document.createElement("a");
      link.href = url;
      link.download = "onlyoffice-probe-" + Date.now() + ".json";
      link.click();
      setTimeout(function () { URL.revokeObjectURL(url); }, 1000);
    };
  };
  plugin.button = function () { plugin.executeCommand("close", ""); };
}());

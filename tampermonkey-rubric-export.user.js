// ==UserScript==
// @name         Moodle Rubric Weight Exporter
// @namespace    https://github.com/raffitch/moodle-gradebook-unifier
// @version      0.1.0
// @description  Adds an export button on Moodle rubric grading pages to download criterion weights as a CSV for the gradebook unifier.
// @match        *://*/grade/edit/grading/form/rubric/*
// @match        *://*/grade/grading/form/rubric/*
// @match        *://*/mod/assign/grade.php*
// @grant        GM_download
// @run-at       document-idle
// ==/UserScript==

(function () {
  "use strict";

  const BUTTON_ID = "moodle-rubric-export-button";

  function waitForRubric(callback) {
    if (document.querySelector(".gradingform_rubric")) {
      callback();
      return;
    }
    const observer = new MutationObserver(() => {
      if (document.querySelector(".gradingform_rubric")) {
        observer.disconnect();
        callback();
      }
    });
    observer.observe(document.documentElement, { childList: true, subtree: true });
  }

  function inferAssignmentName() {
    const heading = document.querySelector(".page-header-headings h1, h1") || document.title;
    const text = heading ? (heading.textContent || heading).trim() : "rubric";
    return text.replace(/\s+/g, " ");
  }

  function parseCriteria() {
    const rows = Array.from(document.querySelectorAll(".gradingform_rubric .criterion"));
    const criteria = [];
    rows.forEach((row, idx) => {
      const labelInput =
        row.querySelector('textarea[name*="[description]"]') ||
        row.querySelector('input[name*="[description]"]') ||
        row.querySelector(".criterionname, .description, .label, .criterionlabel");
      let label = labelInput ? (labelInput.value || labelInput.textContent || "").trim() : "";
      if (!label) {
        label = `Criterion ${idx + 1}`;
      }

      const weightInput = row.querySelector('input[name*="[weight]"]');
      let weight = weightInput ? (weightInput.value || "").trim() : "";
      if (!weight) {
        weight = "0";
      }

      criteria.push({ label, weight });
    });
    return criteria;
  }

  function buildCsv(criteria) {
    if (!criteria.length) return "";
    const lines = [];
    criteria.forEach((c, idx) => {
      const entry = `${c.label} - ${c.weight}%`;
      if (idx === 0) {
        lines.push(entry); // header
      } else {
        lines.push(entry);
      }
    });
    return lines.join("\n");
  }

  function downloadCsv(content, filename) {
    if (typeof GM_download === "function") {
      GM_download({
        url: "data:text/csv;charset=utf-8," + encodeURIComponent(content),
        name: filename,
        saveAs: true,
      });
      return;
    }
    const blob = new Blob([content], { type: "text/csv;charset=utf-8" });
    const link = document.createElement("a");
    link.href = URL.createObjectURL(blob);
    link.download = filename;
    document.body.appendChild(link);
    link.click();
    document.body.removeChild(link);
  }

  function insertButton() {
    if (document.getElementById(BUTTON_ID)) return;
    const host = document.querySelector(".gradingform_rubric") || document.body;

    const container = document.createElement("div");
    container.style.margin = "10px 0";

    const button = document.createElement("button");
    button.id = BUTTON_ID;
    button.type = "button";
    button.textContent = "Export rubric % CSV";
    button.style.padding = "8px 12px";
    button.style.background = "#3b3b3b";
    button.style.color = "white";
    button.style.border = "none";
    button.style.borderRadius = "4px";
    button.style.cursor = "pointer";

    button.addEventListener("click", () => {
      const criteria = parseCriteria();
      if (!criteria.length) {
        alert("No rubric criteria found on this page.");
        return;
      }
      const csv = buildCsv(criteria);
      if (!csv) {
        alert("Unable to build CSV from rubric criteria.");
        return;
      }
      const safeName = inferAssignmentName().replace(/[\\/:*?"<>|]/g, "").trim() || "rubric";
      const filename = `${safeName} - Rubric Percentage.csv`;
      downloadCsv(csv, filename);
    });

    container.appendChild(button);

    const target = document.querySelector(".gradingform_rubric")?.parentElement || document.querySelector("#region-main") || document.body;
    target.insertBefore(container, target.firstChild);
  }

  waitForRubric(insertButton);
})();

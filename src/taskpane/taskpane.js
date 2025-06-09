/*
 * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */

import { all } from "core-js/features/promise";

/* global console, document, Excel, Office */

Office.onReady((info) => {
  if (info.host === Office.HostType.Excel) {
    document.getElementById("add").addEventListener("click",addBoundary)
    document.getElementById("submit").addEventListener("click",setGrades)
  }
});

export async function addBoundary() {
  const div = document.createElement('div');

  div.className = 'row';

  div.innerHTML = `
    <label for="set-mark">Mark:</label>
    <input type="text" name="set-mark" autocomplete="off"/>
    <label for="set-grade">Grade:</label>
    <input type="text"  name="set-grade" autocomplete="off"/>
    <input type="button" value="Remove" onclick="removeBoundary(this)" />`;

  document.getElementById('boundaries').appendChild(div);
}

export async function removeBoundary(input) {
  document.getElementById('boundaries').removeChild(input.parentNode);
}

function collectBoundaries() {
  const boundaries = [];

  document.querySelectorAll('#boundaries .row').forEach(row => {
    const mark = row.querySelector('[name="set-mark"]').value;
    const grade = row.querySelector('[name="set-grade"]').value;

    boundaries.push({ mark, grade });
  });

  for (let i = 0; i < boundaries.length; i++) {
    if (isNaN(boundaries[i].mark)) {
      throw new Error("")
    }
  }

  return boundaries;
}

function collectRanges() {
  const markRange = document.querySelector('[id="mark-range"]').value;
  const gradeRange = document.querySelector('[id="grade-range"]').value;

  return {markRange, gradeRange}
}

function updateMessage(message) {
  const text = message.slice(0, 5);
  if (text === "ERROR") {
    document.getElementById("message").innerHTML = ""
    document.getElementById("error").innerHTML = message
  }
  else {
    document.getElementById("message").innerHTML = message
    document.getElementById("error").innerHTML = ""
  }
}

export async function setGrades() {
  const ranges = collectRanges()
  let boundaries
  try {
    boundaries = collectBoundaries()
  } catch (error) {
    updateMessage("ERROR: Please ensure all boundaries have numerical values")
    return
  }

  Excel.run(async (context) => {
    const sheet = context.workbook.worksheets.getActiveWorksheet()

    let markRange, gradeRange

    try {
      markRange = sheet.getRange(ranges.markRange);
      gradeRange = sheet.getRange(ranges.gradeRange);
    } catch (error) {
      updateMessage("ERROR: Please ensure ranges are only one column each (i.e. B2:B10 and C2:C10)")
      return
    }

    markRange.load("values");
    await context.sync()

    const studentMarks = markRange.values.map(row => row[0])

    let grades
    try {
      grades = studentMarks.map(mark => {
      if (mark === "" || isNaN(mark)) {return "";}

      const numMark = Number(mark);

      for (let i = 0; i < boundaries.length; i++) {
        if (numMark >= boundaries[i].mark) {
          return boundaries[i].grade;
        }
      }
        })
    } catch (error) {
      updateMessage("ERROR: Mark and Grade ranges do not match in length!")
      return
    }
    
    const grades2D = grades.map(grade => [grade]);
    gradeRange.numberFormat = [['@']];
    gradeRange.values = grades2D
    await context.sync()
  })
  updateMessage("Grades set successfully!")
}


window.removeBoundary = removeBoundary;
/*
 * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */

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
    <label for="set-mark">Minimum Mark:</label>
    <input type="text" name="set-mark" />
    <label for="set-grade">Grade:</label>
    <input type="text"  name="set-grade" />
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

  return boundaries;
}

function collectRanges() {
  const markRange = document.querySelector('[id="mark-boundary"]').value;
  const gradeRange = document.querySelector('[id="grade-boundary"]').value;

  return {markRange, gradeRange}
}

export async function setGrades() {
  const boundaries = collectBoundaries()
  for (let i = 0; i < boundaries.length; i++) {
    console.log("Mark: " + boundaries[i].mark + " |  Grade: " + boundaries[i].grade)
  }
  const ranges = collectRanges()
  console.log("Mark Range: " + ranges.markRange + " | Grade Range: " + ranges.gradeRange)
  Excel.run(async (context) => {
    const sheet = context.workbook.worksheets.getActiveWorksheet()

    const markRange = sheet.getRange(ranges.markRange)
    markRange.load("values");
    await context.sync()

    const studentMarks = markRange.values.map(row => row[0])

    const grades = studentMarks.map(mark => {
      if (mark === "" || isNaN(mark)) {return "";}

      const numMark = Number(mark);

      for (let i = 0; i < boundaries.length; i++) {
        if (numMark >= boundaries[i].mark) {
          return boundaries[i].grade;
        }
      }
        })
  

    const gradeRange = sheet.getRange(ranges.gradeRange)
    const grades2D = grades.map(grade => [grade]);
    gradeRange.numberFormat = [['@']];
    gradeRange.values = grades2D
    await context.sync()
  })
  
}

window.removeBoundary = removeBoundary;
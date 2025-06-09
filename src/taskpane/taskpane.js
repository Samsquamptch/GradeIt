/*
 * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */

/* global console, document, Excel, Office */

Office.onReady((info) => {
  if (info.host === Office.HostType.Excel) {
    document.getElementById("add").addEventListener("click",addBoundary)
    document.getElementById("submit").addEventListener("click",setGrades)
    document.getElementById("remove").addEventListener("click",removeBoundary)
  }
});

export async function addBoundary() {
  const div = document.createElement('div');

  div.className = 'row';

  div.innerHTML = `
    <label for="set-mark">Minimum Mark:</label>
    <input type="text" id="set-mark" name="mark-boundary" />
    <label for="set-grade">Grade:</label>
    <input type="text" id="set-grade" name="grade-boundary" />
    <input type="button" value="Remove" onclick="removeBoundary(this)" />`;

  document.getElementById('boundaries').appendChild(div);
}

export async function removeBoundary(input) {
  document.getElementById('boundaries').removeChild(input.parentNode);
}

export async function setGrades() {
  console.log("hello, world")
}

window.removeBoundary = removeBoundary;
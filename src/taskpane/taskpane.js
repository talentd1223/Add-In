/* eslint-disable @typescript-eslint/no-unused-vars */
/* eslint-disable no-undef */
/*
 * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */

/* global document, Office, Word */

// eslint-disable-next-line no-undef
const myobj = require("../../ClaimsLetters.json");
var templateOptions = "";
var genericTextsOptions = "";
var productListOptions = "";
var productTextsOptions = "";
var i;
var proId = "";
var downfilesURI = [];
var downfilesBuf = [];
//var months = [{"jan":0}, {"feb":1}, {"mar":2}, {"apr":3}, {"may":4}, {"jun":5}, {"jul":6}, {"aug":7}, {"sep":8}, {"oct":9}, {"nov":10}, {"dec":11}];
var months = { jan: 0, feb: 2, mar: 3, apr: 4, may: 5, jun: 6, jul: 7, aug: 8, sep: 9, oct: 10, nov: 11, dec: 12 };
//var months = [{"jan":0}, {"feb":1}, {"mar":2}, {"apr":3}, {"may":4}, {"jun":5}, {"jul":6}, {"aug":7}, {"sep":8}, {"oct":9}, {"nov":10}, {"dec":11}];
Office.onReady((info) => {
  if (info.host === Office.HostType.Word) {
    // Determine if the user's version of Office supports all the Office.js APIs that are used in the tutorial.
    if (!Office.context.requirements.isSetSupported("WordApi", "1.3")) {
      // eslint-disable-next-line no-undef
      console.log("Sorry. The tutorial add-in uses Word.js APIs that are not available in your version of Office.");
    }
    //json, docx download;
    initDownSet();

    for (i in myobj.Templates) {
      templateOptions = templateOptions + '<option value="' + i + '">' + myobj.Templates[i].Name + "</option>";
      downfilesURI[i] = myobj.Templates[i].URI;
    }

    for (i in myobj.GenericTexts) {
      genericTextsOptions =
        genericTextsOptions + '<option value="' + i + '">' + myobj.GenericTexts[i].Name + "</option>";
    }

    for (i in myobj.ProductList) {
      if (i == "0") proId = myobj.ProductList[i].Id;
      productListOptions = productListOptions + '<option value="' + i + '">' + myobj.ProductList[i].Name + "</option>";
    }

    for (i in myobj.ProductTexts) {
      if (proId == myobj.ProductTexts[i].Id)
        productTextsOptions =
          productTextsOptions + '<option value="' + i + '">' + myobj.ProductTexts[i].Name + "</option>";
    }

    //context
    document.getElementById("templates").innerHTML = templateOptions;
    document.getElementById("texts").innerHTML = genericTextsOptions;
    document.getElementById("products").innerHTML = productListOptions;
    document.getElementById("product_texts").innerHTML = productTextsOptions;

    // Assign event handlers and other initialization logic.
    document.getElementById("app-body").style.display = "flex";

    //my actions
    document.getElementById("templates").onchange = insertTemplates;
    document.getElementById("texts").onchange = insertGenericTexts;
    document.getElementById("products").onchange = selectProducts;
    document.getElementById("product_texts").onchange = insertProducts;
    document.getElementById("id-datepicker-1").onpointerleave = selectProducts;
  }
});

export async function initDownSet() {
  return Word.run(async (context) => {
    var xhttp = new XMLHttpRequest();
    for (i in downfilesURI) {
      xhttp.open("GET", downfilesURI[i], false);
      xhttp.send();
      downfilesBuf[i] = xhttp.responseText.replace(/^.+,/, "");
      //context.document.body.insertFileFromBase64(downfilesBuf[i], Word.InsertLocation.end);
    }
    await context.sync();
  });
}

function insertProducts() {
  Word.run(function (context) {
    // TODO1: Queue commands to insert a paragraph into the document.
    var docBody = context.document.body;

    docBody.insertParagraph(
      "Name: " + myobj.ProductTexts[document.getElementById("product_texts").value].Name,
      Word.InsertLocation.end
    );
    docBody.insertParagraph(
      " StartDate: " + myobj.ProductTexts[document.getElementById("product_texts").value].StartDate,
      Word.InsertLocation.end
    );
    docBody.insertParagraph(
      " EndDate: " + myobj.ProductTexts[document.getElementById("product_texts").value].EndDate,
      Word.InsertLocation.end
    );
    docBody.insertParagraph(
      " Body: " + myobj.ProductTexts[document.getElementById("product_texts").value].Body,
      Word.InsertLocation.end
    );

    return context.sync();
  }).catch(function (error) {
    // eslint-disable-next-line no-undef
    console.log("Error: " + error);
    if (error instanceof OfficeExtension.Error) {
      console.log("Debug info: " + JSON.stringify(error.debugInfo));
    }
  });
}

function refreshProducts() {
  Word.run(function (context) {
    // TODO1: Queue commands to insert a paragraph into the document.
    productTextsOptions = "";

    for (i in myobj.ProductTexts) {
      var st = myobj.ProductTexts[i].StartDate.split("-");
      var en = myobj.ProductTexts[i].EndDate.split("-");
      var or = document.getElementById("id-textbox-1").value.split("-");

      var stdate = new Date(st[2], months[st[1]], st[0]);
      var endate = new Date(en[2], months[en[1]], en[0]);
      var ordate = stdate;
      if (or.length == 3) ordate = new Date(or[2], months[or[1]], or[0]);

      if (proId == myobj.ProductTexts[i].Id && stdate <= ordate && endate >= ordate)
        productTextsOptions =
          productTextsOptions + '<option value="' + i + '">' + myobj.ProductTexts[i].Name + "</option>";
    }

    document.getElementById("product_texts").innerHTML = productTextsOptions;

    return context.sync();
  }).catch(function (error) {
    // eslint-disable-next-line no-undef
    console.log("Error: " + error);
    if (error instanceof OfficeExtension.Error) {
      console.log("Debug info: " + JSON.stringify(error.debugInfo));
    }
  });
}

function selectProducts() {
  Word.run(function (context) {
    // TODO1: Queue commands to insert a paragraph into the document.

    proId = myobj.ProductList[document.getElementById("products").value].Id;

    refreshProducts();

    return context.sync();
  }).catch(function (error) {
    // eslint-disable-next-line no-undef
    console.log("Error: " + error);
    if (error instanceof OfficeExtension.Error) {
      console.log("Debug info: " + JSON.stringify(error.debugInfo));
    }
  });
}

function insertGenericTexts() {
  Word.run(function (context) {
    // TODO1: Queue commands to insert a paragraph into the document.

    const paragraph = context.document.body.insertParagraph(
      myobj.GenericTexts[document.getElementById("texts").value].Body,
      Word.InsertLocation.end
    );
    //paragraph.font.color = "blue";

    return context.sync();
  }).catch(function (error) {
    // eslint-disable-next-line no-undef
    console.log("Error: " + error);
    if (error instanceof OfficeExtension.Error) {
      console.log("Debug info: " + JSON.stringify(error.debugInfo));
    }
  });
}

export async function insertTemplates() {
  return Word.run(async (context) => {
    var id = document.getElementById("templates").value;

    document.getElementById("invisible").style.visibility = "visible";
    const paragraph = context.document.body.insertFileFromBase64(downfilesBuf[id], Word.InsertLocation.replace);

    await context.sync();
  });
}

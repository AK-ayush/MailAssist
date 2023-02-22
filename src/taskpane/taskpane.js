/*
 * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */

/* global document, Office */

import { callLLMApi } from "../LLM/apiHelper";

var originalBody = "";
var cleanedOriginalBody = "";
var existingSubjects = new Set([]);
var existingSummaries = new Set([]);
var temperature = 0.7;
const firstLineSet = new Set([
  "action items extracted from the mail",
  "action required",
  "actions required",
  "points to extract",
  "key points",
  "action items",
  "key actions",
  "action points",
  "action needed",
  "action",
  "point",
  "points",
  "answer",
  "action items",
  "items",
  "points of email",
  "points to note",
  "actionable points",
  "points extracted",
  "points to be noted",
]);
const maxHighlights = 10;

Office.initialize = function (reason) {
  $(document).ready(function () {
    // Set up ItemChanged event
    Office.context.mailbox.addHandlerAsync(Office.EventType.ItemChanged, itemChanged);

    UpdateTaskPaneUI(Office.context.mailbox.item);
  });
};

function itemChanged(eventArgs) {
  // Update UI based on the new current item
  UpdateTaskPaneUI(Office.context.mailbox.item);
}

// Example implementation
function UpdateTaskPaneUI(item) {
  // Assuming that item is always a read item (instead of a compose item).
  if (item != null) {
    console.log(item.subject);
    let summaryList = document.getElementById("summary-list");
    summaryList.innerHTML = "";
    showLoading();
    displaySummary();
  }
}

Office.onReady((info) => {
  if (info.host === Office.HostType.Outlook) {
    // READ MESSAGE TYPE
    if (Office.context.mailbox.item.displayReplyForm != undefined) {
      let bodyDescriptionText = document.createTextNode("Highlights!");
      let bodyDescriptionDiv = document.getElementById("body-description");
      bodyDescriptionDiv.appendChild(bodyDescriptionText);
      document.getElementById("subjects").style.display = "none";
      document.getElementById("sideload-msg").style.display = "none";
      document.getElementById("more").style.display = "none";
      document.getElementById("format").style.display = "none";
      document.getElementById("reset").style.display = "none";
      //displaySummary();
    } /* COMPOSE MESSAGE TYPE */ else {
      let bodyDescriptionText = document.createTextNode("Suggested Subjects!");
      let bodyDescriptionDiv = document.getElementById("body-description");
      bodyDescriptionDiv.appendChild(bodyDescriptionText);
      document.getElementById("sideload-msg").style.display = "none";
      document.getElementById("app-body").style.display = "flex";
      document.getElementById("summary-list").style.display = "none";
      document.getElementById("more").onclick = run;
      document.getElementById("format").onclick = format;
      document.getElementById("reset").onclick = reset;
      run();
    }
  }
});

// FORMAT MAIL API CALL
export async function callMailFormatter() {
  return await callLLMApi(cleanedOriginalBody, undefined, 1, 2, 1000, 1, 2)
    .then((data) => {
      return data;
    })
    .catch((error) => console.log("error ", error));
}

// FORMAT EVENT
export async function format() {
  var btn = document.getElementById("format");
  btn.disabled = true;
  showLoading();
  // RE-FORMATTING
  if (cleanedOriginalBody.length != 0) {
    callMailFormatter().then((data) => {
      setEmailBody(data[0].text);
    });
  } /* FIRST FORMAT */ else {
    const item = Office.context.mailbox.item;
    item.body.getAsync(Office.CoercionType.Text, (result) => {
      if (result.status === Office.AsyncResultStatus.Succeeded) {
        originalBody = result.value;
        cleanedOriginalBody = cleanEmailBody(originalBody);
        callMailFormatter().then((data) => {
          setEmailBody(data[0].text);
        });
      }
    });
  }
}

// SET MAIL BODY
export async function setEmailBody(emailBody) {
  const item = Office.context.mailbox.item;
  item.body.setAsync(emailBody, (result) => {
    if (result.status !== Office.AsyncResultStatus.Succeeded) {
      return;
    }
    removeLoading();
    var btn = document.getElementById("format");
    btn.disabled = false;
    btn = document.getElementById("reset");
    btn.disabled = false;
  });
}

// RESET EVENT
export async function reset() {
  var btn = document.getElementById("reset");
  btn.disabled = true;
  showLoading();
  setEmailBody(originalBody);
}

// SET MAIL SUBJECT
export async function setSubject() {
  const item = Office.context.mailbox.item;
  item.subject.setAsync(this.innerHTML, (result) => {
    if (result.status !== Office.AsyncResultStatus.Succeeded) {
      return;
    }
  });
}

// SET SUMMARY
export async function displaySummary() {
  const item = Office.context.mailbox.item;
  item.body.getAsync(Office.CoercionType.Text, (result) => {
    if (result.status === Office.AsyncResultStatus.Succeeded) {
      console.log("DISPLAY SUMMARY = " + result.value.toString());
      getSummary(cleanEmailBody(result.value.toString())).then((data) => {
        let summaryArray = data[0].text.split("\n").map((line) =>
          line
            .trim()
            .replace(/^\d+\. /, "")
            .replace(/^\W+/, "")
            .replace(/^(.*\w)[^\w]*$/, "$1")
        );
        for (var i = 0; i < Math.min(summaryArray.length, maxHighlights); i++) {
          if (summaryArray[i] === "" || firstLineSet.has(summaryArray[i].toLowerCase())) {
            continue;
          }
          if (!existingSummaries.has(summaryArray[i].toLowerCase())) {
            existingSummaries.add(summaryArray[i].toLowerCase());
            let summaryListItem = document.createElement("li");
            let summaryText = document.createTextNode(summaryArray[i]);
            let summaryList = document.getElementById("summary-list");
            summaryListItem.className = "listItem-1 pb-2";
            summaryListItem.appendChild(summaryText);
            summaryList.appendChild(summaryListItem);
          }
        }
        console.log(summaryArray);
      });
    }
  });
}

// GET SUMMARY API CALL
export async function getSummary(emailBody) {
  return await callLLMApi(emailBody, undefined, 1, 3, 300, 1.2, 1)
    .then((data) => {
      removeLoading();
      return data;
    })
    .catch((error) => console.log("error ", error));
}

// REMOVE LOADING BUTTON
export async function removeLoading() {
  var div = document.getElementById("loading");
  div.style.display = "none";
}

// SHOW LOADING BUTTON
export async function showLoading() {
  var div = document.getElementById("loading");
  div.style.display = "flex";
}

// GET BODY TO RECOMMEND SUBJECTS
export async function run() {
  const item = Office.context.mailbox.item;
  item.subject.getAsync((result) => {
    showLoading();
    if (result.status === Office.AsyncResultStatus.Succeeded) {
      if (!result.value.toString().toLowerCase().startsWith("re:")) {
        var btn = document.getElementById("more");
        btn.disabled = true;
        var contentOfSubject;
        item.getSelectedDataAsync(Office.CoercionType.Text, (asyncResult) => {
          if (asyncResult.status === Office.AsyncResultStatus.Succeeded) {
            contentOfSubject = asyncResult.value.data;
          }
          if (contentOfSubject.length === 0) {
            item.body.getAsync(Office.CoercionType.Text, (result) => {
              if (result.status === Office.AsyncResultStatus.Succeeded) {
                contentOfSubject = result.value.toString();
                recommendSubjects(contentOfSubject).then((data) => {
                  removeLoading();
                  displaySubjects(data);
                });
              }
            });
          } else {
            recommendSubjects(contentOfSubject).then((data) => {
              removeLoading();
              displaySubjects(data);
            });
          }
        });
      }
    }
  });
}

// DISPLAY SUBJECTS
export async function displaySubjects(data) {
  data.forEach((subjectData) => {
    const arr = subjectData.text.split(":");
    const subject =
      arr[0].trim().includes("Subject") || arr[0].trim().includes("subject") ? arr[1].trim() : subjectData.text.trim();
    if (subject != "") {
      if (!existingSubjects.has(subject.toLowerCase())) {
        existingSubjects.add(subject.toLowerCase());
        var div = document.createElement("div");
        var t = document.createTextNode(subject);
        div.style.cursor = "pointer";
        div.appendChild(t);
        div.className = "button border border-top-0 border-right-0 m-0";
        div.id = "button-1";
        div.addEventListener("click", setSubject);
        document.getElementById("subjects").appendChild(div);
      }
    }
  });
  temperature += 0.05;
  var btn = document.getElementById("more");
  btn.disabled = false;
}

// CLEAN EMAIL BODY
function cleanEmailBody(emailBody) {
  emailBody = emailBody.replace(/^(From:|Sent:|To:|Cc:|Subject:).*$\n?/gm, "");
  const lines = emailBody.split("\n");

  const greetings = [
    "hello",
    "hi",
    "hey",
    "greetings",
    "best regards",
    "regards",
    "thank you",
    "sincerely",
    "yours sincerely",
  ];

  // Remove the first line if it's a greeting
  if (lines[0].toLowerCase().match(new RegExp(`^(${greetings.join("|")}),?$`, "i"))) {
    lines.shift();
  }

  // Remove the last two lines if they're a greeting
  if (
    lines.length > 2 &&
    lines[lines.length - 2].toLowerCase().match(new RegExp(`^(${greetings.join("|")}),?$`, "i")) &&
    lines[lines.length - 1].match(/^[A-Z][a-z]+$/)
  ) {
    lines.pop();
    lines.pop();
  }

  const emailBodyWithoutGreeting = lines.join("\n");
  const teamsMeetingText = /Microsoft Teams meeting[\s\S]*Meeting options/g;
  const emailBodyWithoutMeetingInfo = emailBodyWithoutGreeting.replace(teamsMeetingText, "");
  const emailText = emailBodyWithoutMeetingInfo.replace(
    /(<([^>]+)>)|(\b(https?|ftp|file):\/\/[-A-Z0-9+&@#\/%?=~_|!:,.;]*[-A-Z0-9+&@#\/%=~_|])/gi,
    ""
  );
  const emailTextWithExtraNewLines = emailText.replace(new RegExp(`\\b(${greetings.join("|")})\\b`, "ig"), "");
  const cleanedBody = emailTextWithExtraNewLines.replace(/\s{2,}/g, " ");
  console.log("Cleaned EMAIL: ", cleanedBody);
  return cleanedBody;
}

// RECOMMEND SUBJECTS API CALL
export async function recommendSubjects(emailBody) {
  const emailTextWithoutGreetingWords = cleanEmailBody(emailBody);

  return await callLLMApi(
    emailTextWithoutGreetingWords,
    undefined,
    undefined,
    undefined,
    undefined,
    temperature
  )
    .then((data) => {
      return data;
    })
    .catch((error) => console.log("error ", error));
}

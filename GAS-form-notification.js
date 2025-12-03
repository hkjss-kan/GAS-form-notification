function sendEmailOnFormSubmit(e) {
  let formName = e.source.getTitle();
  let sendToEmailAddressArray = [
    '<RECIPIENT_EMAIL_ADDRESS>',
  ];
  let emailSubject = "[" + formName + "] Response received";
  let emailHtmlBodyText = "[" + formName + "] Response received<br><br>";

  const formResponse = e.response;
  const itemResponses = formResponse.getItemResponses();
  for (let i = 0; i < itemResponses.length; i++) {
    const itemResponse = itemResponses[i];
    emailHtmlBodyText += itemResponse.getItem().getTitle() +
      ": " +
      itemResponse.getResponse() +
      "<br>";
  }

  emailHtmlBodyText += "<br>All respones tally: <br>";
  let tallyResult = tallyFormResponses(e.source);
  tallyResult.forEach(q => {
    emailHtmlBodyText += "<br>" + q.question + ":<br>";
      // Sort answers by count descending
      const answers = Object.entries(q).filter(([k]) => k !== "question");
      answers.sort((a, b) => b[1] - a[1]);

      answers.forEach(([answer, count]) => {
        emailHtmlBodyText += answer + ": " + count + "<br>";
      });
    });
  
  const sendToEmailAddresses = sendToEmailAddressArray.join(',');
  MailApp.sendEmail(
    { 
      to: sendToEmailAddresses, 
      subject: emailSubject, 
      htmlBody: emailHtmlBodyText,
    }
  );
}

function tallyFormResponses(f) {
  const form = f;
  const responses = form.getResponses();  // All submitted responses
  const items = form.getItems();          // All questions in the form
  
  // This will hold our final result
  let tallyResult = [];

  // Loop through each question (Form Item)
  items.forEach(item => {
    const title = item.getTitle().trim();
    if (!title) return; // Skip items with no title (like sections, images)

    const itemType = item.getType();
    let answerCounts = {};

    // Only process question types that have selectable answers
    if (
      itemType === FormApp.ItemType.MULTIPLE_CHOICE ||
      itemType === FormApp.ItemType.LIST ||
      itemType === FormApp.ItemType.CHECKBOX ||
      itemType === FormApp.ItemType.DROPDOWN
    ) {
      // For choice-based questions
      responses.forEach(response => {
        const itemResponse = response.getResponseForItem(item);
        if (itemResponse) {
          const answers = itemResponse.getResponse(); // Can be string or array (for checkboxes)

          if (Array.isArray(answers)) {
            // Checkbox: multiple answers possible
            answers.forEach(ans => {
              const answerText = ans.trim();
              answerCounts[answerText] = (answerCounts[answerText] || 0) + 1;
            });
          } else {
            // Single choice
            const answerText = answers ? answers.toString().trim() : "(no response)";
            answerCounts[answerText] = (answerCounts[answerText] || 0) + 1;
          }
        } else {
          // No response given for this question
          answerCounts["(no response)"] = (answerCounts["(no response)"] || 0) + 1;
        }
      });

      // Build the final object for this question
      const questionResult = {
        question: title
      };
      Object.assign(questionResult, answerCounts); // Merge counts into the object
      tallyResult.push(questionResult);
    } 
  });

  return tallyResult;
}

// LIMITATIONS:
// 1. The script only works with Google Docs for content
// 2. The script only works with Google Sheets for data
// 3. The script only works with Gmail for sending emails
// 4. Cannot accomodate for special formatting in the subject line
// 5. Cannot accomodate for text size changes in subject or body
// 6. Cannot accomodate for bullet points or numbered lists
// 7. Cannot accomodate for tables
// 8. Cannot accomodate for images
// 9. Cannot accomodate for words that are partially bolded,
// italicized, underlined, highlighted, or colored. You can attempt this
// but expect some bugs to come in your email draft.

// REMEMBER:
// This script only CREATES drafts of emails. You must go to your drafts folder in Gmail
// to send the emails. The script will not send the emails for you.
// Don't forget to select the createOutreachEmails function in the dropdown menu
// in the Apps Script editor before running the script.
// Also change the DOC_ID, NAME_COL, and EMAIL_COL to match the Google Doc ID, the column name
// for the principal's last name, and the column name for the principal's email respectively.

// The ID of the Google Doc that contains the email content
const DOC_ID = '1QxxoIzQDU6B12u6484QyDLYTaFNW8yr2nwQf54pU7G4'
// The name of the column in the Google Sheet that contains the principal's last name
const NAME_COL = "Principal Last Name"
// The name of the column in the Google Sheet that contains the principal's email
const EMAIL_COL = "Principal's email"
// The name o fht ecolumn in the Google Sheet that contains the principal's school
const SCHOOL_COL = "School Name"
// The number of times something was logged to the console (for debugging)
var LOG_COUNT = 0
const LOG_LIMIT = 10

function createOutreachEmails() {
	const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
	const data = sheet.getDataRange().getValues();
	const headers = data.shift(); // remove the header row for processing

	// identify the index of each column based on the header
	const nameColIndex = headers.indexOf(NAME_COL);
	const emailColIndex = headers.indexOf(EMAIL_COL);
	const schoolColIndex = headers.indexOf(SCHOOL_COL);

	// check if the columns exist
	if (nameColIndex === -1 || emailColIndex === -1) {
		console.log(
			"One or more required columns are missing."
		);
		return;
	}
    // get the Google Doc content
    const doc = DocumentApp.openById(DOC_ID);
    const paragraphs = doc.getBody().getParagraphs();
    const subject = paragraphs[0].getText();
    const body = paragraphs.slice(1).map(p => convertParagraphToHtml(p)).join('<br>');
    data.forEach((row) => {
        const LastName = row[nameColIndex];
        const Email = row[emailColIndex];
		const School = row[schoolColIndex];
		console.log(
			`Processing row: LastName='${LastName}', Email='${Email}, School=${School}'`
		);

		if (!LastName || !Email) {
            console.log(
                `Skipping row due to missing data: LastName='${LastName}', Email='${Email}, School=${School}'`
            );
            return; // skip this iteration
        }
		// For any future users: you can change the subject and body of the email here. For any instances in
		// which you want to include the value of a variable, use ${variableName}. For example, the subject
		// below includes the value of the LastName variable for the principal's last name. Also, if you want
		// to include a single quote, double quote, backslash, or any other special character in the email,
		// you must escape it with a backslash. For example, to include a single quote in the subject, you
		// would write \'. For more information on string literals in JavaScript, see the following link:
		// https://developer.mozilla.org/en-US/docs/Web/JavaScript/Reference/Global_Objects/String
		
        const personalizedSubject = subject.replace(/\${LastName}/g, LastName).replace(/\${Email}/g, Email).replace(/\${School}/g, School);
        const personalizedBody = body.replace(/\${LastName}/g, LastName).replace(/\${Email}/g, Email).replace(/\${School}/g, School);

        // create the email draft
        GmailApp.createDraft(Email, personalizedSubject, '', {htmlBody: personalizedBody});

        console.log(
            `Draft created for ${LastName} (${Email}) with subject: '${personalizedSubject}'`
        );
    });

    console.log("Outreach email draft creation process completed.");
}

function convertParagraphToHtml(paragraph) {
    let html = '';

    for (let i = 0; i < paragraph.getNumChildren(); i++) {
        const element = paragraph.getChild(i);
        if (element.getType() === DocumentApp.ElementType.TEXT) {
            const text = element.asText();
            let wordStart = 0;
            let wordEnd = 0;

            while (wordEnd < text.getText().length) {
                if (text.getText().charAt(wordEnd).match(/\s/)) {
                    let word = text.getText().substring(wordStart, wordEnd);
                    let wordHtml = word;
                    const attributes = text.getAttributes(wordStart);

                    if (attributes[DocumentApp.Attribute.BOLD]) {
                        wordHtml = `<b>${wordHtml}</b>`;
                    }
                    if (attributes[DocumentApp.Attribute.ITALIC]) {
                        wordHtml = `<i>${wordHtml}</i>`;
                    }
                    if (attributes[DocumentApp.Attribute.UNDERLINE]) {
                        wordHtml = `<u>${wordHtml}</u>`;
                    }
                    if (attributes[DocumentApp.Attribute.BACKGROUND_COLOR] && attributes[DocumentApp.Attribute.BACKGROUND_COLOR] !== '#ffffff' && attributes[DocumentApp.Attribute.BACKGROUND_COLOR] !== '#000000') {
                        wordHtml = `<span style="background-color:${attributes[DocumentApp.Attribute.BACKGROUND_COLOR]}">${wordHtml}</span>`;
                    }
                    if (attributes[DocumentApp.Attribute.FOREGROUND_COLOR] && attributes[DocumentApp.Attribute.FOREGROUND_COLOR] !== '#000000') {
                        wordHtml = `<span style="color:${attributes[DocumentApp.Attribute.FOREGROUND_COLOR]}">${wordHtml}</span>`;
                    }
                    html += wordHtml + text.getText().charAt(wordEnd);
                    wordStart = wordEnd + 1;
                }
                wordEnd++;
            }

            // Process the last word if there is no trailing space
            if (wordStart < text.getText().length) {
                let word = text.getText().substring(wordStart);
                let wordHtml = word;
                const attributes = text.getAttributes(wordStart);

                if (attributes[DocumentApp.Attribute.BOLD]) {
                    wordHtml = `<b>${wordHtml}</b>`;
                }
                if (attributes[DocumentApp.Attribute.ITALIC]) {
                    wordHtml = `<i>${wordHtml}</i>`;
                }
                if (attributes[DocumentApp.Attribute.UNDERLINE]) {
                    wordHtml = `<u>${wordHtml}</u>`;
                }
                if (attributes[DocumentApp.Attribute.BACKGROUND_COLOR] && attributes[DocumentApp.Attribute.BACKGROUND_COLOR] !== '#ffffff' && attributes[DocumentApp.Attribute.BACKGROUND_COLOR] !== '#000000') {
                    wordHtml = `<span style="background-color:${attributes[DocumentApp.Attribute.BACKGROUND_COLOR]}">${wordHtml}</span>`;
                }
                if (attributes[DocumentApp.Attribute.FOREGROUND_COLOR] && attributes[DocumentApp.Attribute.FOREGROUND_COLOR] !== '#000000') {
                    wordHtml = `<span style="color:${attributes[DocumentApp.Attribute.FOREGROUND_COLOR]}">${wordHtml}</span>`;
                }
                html += wordHtml;
            }
        }
    }

    return html;
}
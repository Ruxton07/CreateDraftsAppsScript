# createDraftsAppsScript project

This Google Apps Script project creates email drafts in Gmail using data from a Google Sheet and content from a Google Doc. The script supports various text formatting options and allows for the inclusion of special characters and variables in the email content.

The script will not send any emails, but can only create drafts. If you have a Gmail signature, then do not include it in the Google doc (Gmail automatically puts your signature in any drafts).

The first paragraph is reserved for the subject line, and any subsequent paragraphs in the Google doc will be added into the body section.

## Features

- Supports bold, underline, italics, text coloring, and text highlighting in the email body.
- Allows the inclusion of special characters in the email subject and body.
- Uses variables from the Google Sheet data in the email content.

## Limitations

1. The script only works with Google Docs for content.
2. The script only works with Google Sheets for data.
3. The script only works with Gmail for sending emails.
4. Cannot accommodate special formatting in the subject line.
5. Cannot accommodate text size changes in the subject or body.
6. Cannot accommodate bullet points or numbered lists.
7. Cannot accommodate tables.
8. Cannot accommodate images.
9. Cannot accommodate words that are partially bolded, italicized, underlined, highlighted, or colored. You can attempt this but expect some bugs to come in your email draft.

## How to Use

1. **Set Up the Script:**
   - Open the Google Apps Script editor.
   - Create a new project and copy the contents of `Code.js` into the script editor.
   - Save the project.

2. **Configure the Script:**
   - Update the `DOC_ID` constant with the ID of your Google Doc that contains the email content.
   - Update the `NAME_COL`, `EMAIL_COL`, and `SCHOOL_COL` constants to match the column names in your Google Sheet for the principal's last name, email, and school name, respectively.

3. **Run the Script:**
   - In the Apps Script editor, select the `createOutreachEmails` function from the dropdown menu.
   - Click the run button to execute the script.

4. **Check Your Drafts:**
   - The script will create email drafts in your Gmail account.
   - Go to your Gmail drafts folder to review and send the emails.

## Including Variables in Email Content

You can include variables from the Google Sheet data in the email subject and body by using the following placeholders:
- `${LastName}`: The principal's last name.
- `${Email}`: The principal's email.
- `${School}`: The principal's school name.

For example, if you want to include the principal's last name in the subject, you can write:

```
Welcome, ${LastName}!
```

## Special Characters

To include special characters such as single quotes, double quotes, or backslashes in the email content, you must escape them with a backslash. For example, to include a single quote in the subject, you would write:

```
Welcome to the school's event, it's going to be great!😊
```

The only exception to this is that you risk breaking the script if you try to add any formatting to an emoji or other unicode character.


For more information on string literals in JavaScript, see the following link:
[JavaScript String Literals](https://developer.mozilla.org/en-US/docs/Web/JavaScript/Reference/Global_Objects/String)

## Example

Here is an example of how the email content might look in the Google Doc:

```
Subject: Welcome, ${LastName}!

Dear ${LastName},

We are excited to have you at ${School}. Please find the details below:

- Event Date: January 1, 2023
- Location: Main Auditorium
- Best regards, The Team
```

This will create an email draft with the subject "Welcome, [LastName]!" and the body with the appropriate formatting and variable replacements.

## Conclusion

This script helps automate the creation of personalized email drafts in Gmail using data from a Google Sheet and content from a Google Doc. It supports various text formatting options and allows for the inclusion of special characters and variables in the email content.
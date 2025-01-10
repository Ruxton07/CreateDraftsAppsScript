function createBirthdayEmailDrafts() {
	const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
	const data = sheet.getDataRange().getValues();
	const headers = data.shift(); // remove the header row for processing

	// identify the index of each column based on the header
	const nameColIndex = headers.indexOf("Principal Last Name");
	const emailColIndex = headers.indexOf("Principal's email");

	// check if the columns exist
	if (nameColIndex === -1 || emailColIndex === -1) {
		console.log(
			"One or more required columns ('Principal Last Name', 'Principal's email') are missing."
		);
		return;
	}

	data.forEach((row) => {
		const LastName = row[nameColIndex];
		const Email = row[emailColIndex];

		if (!LastName || !Email) {
			console.log(
				`Skipping row due to missing data: LastName='${LastName}', Email='${Email}'`
			);
			return; // skip this iteration
		}

		const subject = `happy BDAY here is test characters ' * " \\  ) ! @ # $ % ^ & *  - + = _(, ${LastName}!`;
		const body = `Dear ${LastName},\n\nyour birthday is literally my favorite!\n\nBest wishes,\nLASA 418`;

		// create the email draft
		GmailApp.createDraft(Email, subject, body);

		console.log(
			`Draft created for ${LastName} (${Email}) with subject: '${subject}'`
		);
	});

	console.log("Birthday email draft creation process completed.");
}

const { google } = require('googleapis');
const { authenticate } = require('@google-cloud/local-auth');
const fs = require('fs').promises;
const path = require('path');

const authService = async (serviceType) => {
	try {
		const credentials = await fs.readFile('credentials.json');

		// scopes for GMAIL API - reading, sending, labelling
		const scopes = [
			'https://www.googleapis.com/auth/gmail.readonly',
			'https://www.googleapis.com/auth/gmail.send',
			'https://www.googleapis.com/auth/gmail.labels',
			'https://mail.google.com/',
		];

		if (serviceType === 'outlook') {
		}

		if (serviceType === 'google') {
			// authenticating and calling the gmail API
			const auth = await authenticate({
				keyfilePath: path.join(process.cwd(), 'credentials.json'),
				scopes,
			});

			const gmail = google.gmail({ version: 'v1', auth });

			const response = await gmail.users.labels.list({
				userId: 'me',
			});

			// use a label name using AI;
			const LABEL_NAME = 'TEST_LABEL';

			return { auth, gmail, response, LABEL_NAME };
		}
	} catch (error) {
		console.log(error);
	}
};

module.exports = { authService };

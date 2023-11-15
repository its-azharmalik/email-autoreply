const { google } = require('googleapis');

// function get unreplied messages
const getUnrepliedMessages = async (auth) => {
	const gmail = google.gmail({ version: 'v1', auth });
	const res = await gmail.users.messages.list({
		userId: 'me',
		q: '-in:chats -from:me -has:userlabels',
	});
	return res.data.messages || [];
};

// function send reply to the messages
const sendReply = async (auth, message) => {
	const gmail = google.gmail({ version: 'v1', auth });
	const messageDetails = await gmail.users.messages.get({
		userId: 'me',
		id: message.id,
		format: 'metadata',
		metadataHeaders: ['To', 'Subject', 'From'],
	});
	console.log(messageDetails.data.payload);

	const subject = messageDetails.data.payload.headers.find(
		(header) => header.name === 'Subject'
	).value;
	const from = messageDetails.data.payload.headers.find(
		(header) => header.name === 'From'
	).value;

	const replyTo = from.match(/<(.*)>/)[1];
	const replySubject = subject.startsWith('Re: ') ? subject : `Re: ${subject}`;
	const replyBody =
		'Hello! \n\nThank you for your email. \n\nI am currently on vacation and will reply to your email as soon as possible. \n\nBest Regards, \n\nAzhar Mahmood';

	const messageFormat = [
		`From: me`,
		`To: ${replyTo}`,
		`Subject: ${replySubject}`,
		`In-Reply-To: ${message.id}`,
		`References: ${message.id}`,
		'',
		replyBody,
	].join('\n');

	const encodedMessage = Buffer.from(messageFormat)
		.toString('base64')
		.replace(/\+/g, '-')
		.replace(/\//g, '_')
		.replace(/=+$/, '');

	const res = await gmail.users.messages.send({
		userId: 'me',
		requestBody: {
			raw: encodedMessage,
		},
	});
	return res;
};

// function to  create a label
const createLabel = async (auth, LABEL_NAME) => {
	const gmail = google.gmail({ version: 'v1', auth });
	try {
		const res = await gmail.users.labels.create({
			userId: 'me',
			requestBody: {
				name: LABEL_NAME,
				labelListVisibility: 'labelShow',
				messageListVisibility: 'show',
			},
		});
		return res.data;
	} catch (error) {
		if (error.code == 409) {
			// label already exists
			const res = await gmail.users.labels.list({
				userId: 'me',
			});

			const label = res.data.labels.find((label) => label.name === LABEL_NAME);
			return label;
		} else throw error;
	}
};

// function add the label to the message and move it to the label folder
const addLabel = async (auth, message, labelId) => {
	const gmail = google.gmail({ version: 'v1', auth });
	const res = await gmail.users.messages.modify({
		userId: 'me',
		id: message.id,
		requestBody: {
			addLabelIds: [labelId],
			removeLabelIds: ['INBOX'],
		},
	});
	return res.data;
};

module.exports = {
	addLabel,
	createLabel,
	getUnrepliedMessages,
	sendReply,
};

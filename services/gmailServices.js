const {
	createLabel,
	getUnrepliedMessages,
	sendReply,
	addLabel,
} = require('./helpers/gmailHelpers');

const gmailService = async (auth, LABEL_NAME) => {
	// Create a Label
	const label = await createLabel(auth, LABEL_NAME);

	// setup a cronjob in intervals to repeat the following steps
	setInterval(async () => {
		// get the unreplied messages
		const messages = await getUnrepliedMessages(auth);
		console.log(messages.length, ' - Messages Found');

		for (const message of messages) {
			// send reply to the message
			const reply = await sendReply(auth, message);
			console.log(`Replied to ${message.id} - `, reply.data);

			// add the label to the message and move it to the label folder
			const addedLabel = await addLabel(auth, message, label.id);
			console.log(`Added label to ${message.id}`);
		}
	}, 60000); // check every minute for new mails
};

module.exports = { gmailService };

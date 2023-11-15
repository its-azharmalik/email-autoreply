const express = require('express');
const colors = require('colors');

// Services
const { gmailService } = require('./services/gmailService');
const { authService } = require('./services/authService');

// Server Setup
const PORT = 8080;
const app = express();

app.get('/', async (req, res) => {
	const { auth, LABEL_NAME } = await authService();
	await gmailService(auth, LABEL_NAME).catch(console.error);
	if (auth?.credentials.access_token === undefined) {
		res.send('There is some error please retry login');
	} else res.send('You have successfully subscribed to our services');
});

app.listen(PORT, () => {
	console.log(`App running on PORT ${`${PORT}`.bold.yellow}`);
});

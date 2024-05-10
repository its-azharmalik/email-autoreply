const express = require('express');
const colors = require('colors');
const cors = require('cors');
const axios = require('axios');
const passport = require('passport');
const OIDCStrategy = require('passport-azure-ad').OIDCStrategy;
// Client Setup OutLook
const { Client } = require('@microsoft/microsoft-graph-client');
const {
	PublicClientApplication,
	ConfidentialClientApplication,
} = require('@azure/msal-node');

// Services
const { gmailService } = require('./services/gmailService');
const { authService } = require('./services/authService');
const session = require('express-session');

// Server Setup
const PORT = 8080;
const app = express();

// Middlewares
app.use(cors());
app.use(express.json());
app.use(express.urlencoded({ extended: true }));
app.use(
	session({
		secret: 'any_secret_key',
		resave: false,
		saveUninitialized: false,
	})
);

passport.use(
	new OIDCStrategy(
		{
			identityMetadata: `https://login.microsoftonline.com/${MICROSOFT_TENANT_ID}/v2.0/.well-known/openid-configuration`,
			clientID: MICROSOFT_CLIENT_ID,
			clientSecret: MICROSOFT_CLIENT_SECRET,
			responseType: 'code',
			responseMode: 'form_post',
			redirectUrl: 'https://4dc4-117-239-210-99.ngrok-free.app/auth/callback', // Replace with HTTPS URL
			passReqToCallback: true,
		},
		function (req, iss, sub, profile, accessToken, refreshToken, done) {
			// Store accessToken and other user details if needed
			return done(null, { accessToken: accessToken });
		}
	)
);

// Use passport session
app.use(passport.initialize());
app.use(passport.session());

// Serialize and deserialize user
passport.serializeUser(function (user, done) {
	done(null, user);
});

passport.deserializeUser(function (user, done) {
	done(null, user);
});

// Routes for authentication
app.get(
	'/auth/microsoft',
	passport.authenticate('azuread-openidconnect', {
		scope: ['openid', 'profile', 'User.Read', 'Mail.Read'],
	})
);

app.post(
	'/auth/callback',
	passport.authenticate('azuread-openidconnect', {
		failureRedirect: '/auth/microsoft',
	}),
	function (req, res) {
		// Successful authentication, redirect to fetch emails
		res.redirect('/fetchEmails');
	}
);

// Middleware to ensure user is authenticated
function ensureAuthenticated(req, res, next) {
	if (req.isAuthenticated()) {
		return next();
	}
	res.redirect('/auth/microsoft');
}

// Route to fetch emails after authentication
app.get('/fetchEmails', ensureAuthenticated, async function (req, res) {
	try {
		// Fetch emails using Microsoft Graph API
		const emailsResponse = await axios.get(
			'https://graph.microsoft.com/v1.0/me/messages',
			{
				headers: {
					Authorization: `Bearer ${req.user.accessToken}`,
				},
			}
		);
		// Extract relevant email data
		const emails = emailsResponse.data.value.map((email) => ({
			subject: email.subject,
			from: email.sender.emailAddress.address,
			body: email.bodyPreview,
		}));
		// Send email data as response
		res.json(emails);
	} catch (error) {
		res.status(500).json({ error: 'Failed to fetch emails' });
	}
});

const scopes = ['https://graph.microsoft.com/.default'];
const msalConfig = {
	auth: {
		clientId,
		authority: `https://login.microsoftonline.com/${tenantId}`,
		redirectUri,
		clientSecret,
	},
};

const pca = new PublicClientApplication(msalConfig);

const ccaConfig = {
	auth: {
		clientId,
		authority: `https://login.microsoftonline.com/${tenantId}`,
		clientSecret,
	},
};

const cca = new ConfidentialClientApplication(ccaConfig);

app.get('/google', async (req, res) => {
	const { auth, LABEL_NAME } = await authService('google');
	await gmailService(auth, LABEL_NAME).catch(console.error);
	if (auth?.credentials.access_token === undefined) {
		res.send('There is some error please retry login');
	} else res.send('You have successfully subscribed to our services');
});

app.get('/outlook', async (req, res) => {
	// const { auth, LABEL_NAME } = await authService('outlook');

	// run the outlook service
	try {
		const authCodeUrlParameters = {
			scopes,
			redirectUri,
		};
		const response = await pca.getAuthCodeUrl(authCodeUrlParameters);
		res.redirect(response);
	} catch (error) {
		console.log('Outlook', error);
	}

	// check for outllok token and give the response accordingly
	// if (auth?.credentials.access_token === undefined) {
	// 	res.send('There is some error please retry login');
	// } else res.send('You have successfully subscribed to our services');
});

app.get('/', async (req, res) => {
	try {
		const tokenRequest = {
			code: req.query.code,
			scopes,
			redirectUri,
			clientSecret,
		};
		console.log(tokenRequest);
		const resp = await pca.acquireTokenByCode(tokenRequest);

		req.session.accessToken = resp.accessToken;
		console.log(resp);
		req.session.userIdAzure = resp.idTokenClaims.oid;
		res.redirect('/access-token');
		// use this token to call the outlook API and fetch the emails
	} catch (error) {
		console.log('Redirect', error);
	}
});

app.get('/access-token', async (req, res) => {
	try {
		const tokenRequest = {
			scopes,
			clientSecret: clientSecret,
		};
		const { accessToken } = await cca.acquireTokenByClientCredential(
			tokenRequest
		);
		req.session.clientAccessToken = accessToken;
		res.send('Access token acquired successfully!');
	} catch (error) {
		console.log('Access Token', error);
	}
});

app.get('/emails', async (req, res) => {
	try {
		const clientAccessToken = req.session.clientAccessToken;
		const userAccessToken = req.session.accessToken;
		const userId = req.session.userIdAzure;
		console.log({ userAccessToken, clientAccessToken, userId });
		if (!clientAccessToken || !userAccessToken) {
			res.send('Please login first');
		}

		const emailsResponse = await axios.get(
			`https://graph.microsoft.com/v1.0/${userId}/messages`,
			{
				headers: {
					Authorization: `Bearer ${clientAccessToken}`,
					'Content-Type': 'application/json',
				},
			}
		);
		res.send({ emails: emailsResponse });
	} catch (error) {
		console.log('Emails', error.response.data);
		res.send(error?.response.data?.error);
	}
});

app.listen(PORT, () => {
	console.log(`App running on PORT ${`${PORT}`.bold.yellow}`);
});

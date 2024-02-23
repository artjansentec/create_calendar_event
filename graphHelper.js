// Copyright (c) Microsoft Corporation. All rights reserved.
// Licensed under the MIT license.

// <UserAuthConfigSnippet>
// require('./fetchFile');
const azure = require('@azure/identity');
const graph = require('@microsoft/microsoft-graph-client');
const authProviders = require('@microsoft/microsoft-graph-client/authProviders/azureTokenCredentials');

let _settings = undefined;
let _deviceCodeCredential = undefined;
let _userClient = undefined;

function initializeGraphForUserAuth(settings, deviceCodePrompt) {
  // Ensure settings isn't null
  if (!settings) {
    throw new Error('Settings cannot be undefined');
  }

  _settings = settings;

  _deviceCodeCredential = new azure.DeviceCodeCredential({
    clientId: settings.clientId,
    tenantId: settings.tenantId,
    userPromptCallback: deviceCodePrompt
  });

  const authProvider = new authProviders.TokenCredentialAuthenticationProvider(
    _deviceCodeCredential, {
      scopes: settings.graphUserScopes
    });

  _userClient = graph.Client.initWithMiddleware({
    authProvider: authProvider
  });
}
module.exports.initializeGraphForUserAuth = initializeGraphForUserAuth;
// </UserAuthConfigSnippet>

// <GetUserTokenSnippet>
async function getUserTokenAsync() {
  // Ensure credential isn't undefined
  if (!_deviceCodeCredential) {
    throw new Error('Graph has not been initialized for user auth');
  }

  // Ensure scopes isn't undefined
  if (!_settings?.graphUserScopes) {
    throw new Error('Setting "scopes" cannot be undefined');
  }

  // Request token with given scopes
  const response = await _deviceCodeCredential.getToken(_settings?.graphUserScopes);
  return response.token;
}
module.exports.getUserTokenAsync = getUserTokenAsync;
// </GetUserTokenSnippet>

// <GetUserSnippet>
async function getUserAsync() {
  // Ensure client isn't undefined
  if (!_userClient) {
    throw new Error('Graph has not been initialized for user auth');
  }

  return _userClient.api('/me')
    // Only request specific properties
    .select(['displayName', 'mail', 'userPrincipalName'])
    .get();
}
module.exports.getUserAsync = getUserAsync;
// </GetUserSnippet>

// <GetInboxSnippet>
async function getInboxAsync() {
  // Ensure client isn't undefined
  if (!_userClient) {
    throw new Error('Graph has not been initialized for user auth');
  }

  return _userClient.api('/me/mailFolders/inbox/messages')
    .select(['from', 'isRead', 'receivedDateTime', 'subject'])
    .top(25)
    .orderby('receivedDateTime DESC')
    .get();
}
module.exports.getInboxAsync = getInboxAsync;
// </GetInboxSnippet>

// <SendMailSnippet>
async function sendEventAsync(subject, body, recipient) {
  // Ensure client isn't undefined
  if (!_userClient) {
    throw new Error('Graph has not been initialized for user auth');
  }
  // Create a new message
  const event = {
    subject: 'Jorge',
    body: {
      contentType: 'HTML',
      content: 'Pelo amor de Deus jorge e matheus'
    },
    start: {
        dateTime: '2024-02-14T08:07:00',
        timeZone: 'E. South America Standard Time'
    },
    end: {
        dateTime: '2024-02-14T08:10:00',
        timeZone: 'E. South America Standard Time'
    },
    location: {
        displayName: 'Casa do Tutu'
    },
    attendees: [
      {
        emailAddress: {
          address: 'leonardo3.botrel@aquila.com.br',
          name: 'Leonardo'
        }
      },
      {
        emailAddress: {
          address: 'pedrinho.patolouco@polonorte.com.br',
          name: 'Pato Donald'
        }
      }
    ]
  };
  
  return _userClient.api('/me/events')
	.post(event);

}
_userClient
module.exports.sendEventAsync = sendEventAsync;
// </SendMailSnippet>

// <MakeGraphCallSnippet>
// This function serves as a playground for testing Graph snippets
// or other code
async function makeGraphCallAsync() {
  // INSERT YOUR CODE HERE
}
module.exports.makeGraphCallAsync = makeGraphCallAsync;
// </MakeGraphCallSnippet>

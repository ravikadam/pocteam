require('isomorphic-fetch');
const azure = require('@azure/identity');
const graph = require('@microsoft/microsoft-graph-client');
const authProviders =
  require('@microsoft/microsoft-graph-client/authProviders/azureTokenCredentials');

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
    tenantId: settings.authTenant,
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

async function getUserTokenAsync() {
  // Ensure credential isn't undefined
  if (!_deviceCodeCredential) {
    throw new Error('Graph has not been initialized for user auth');
  }

  // Ensure scopes isn't undefined
  if (!_settings.graphUserScopes) {
    throw new Error('Setting "scopes" cannot be undefined');
  }

  // Request token with given scopes
  const response = await _deviceCodeCredential.getToken(_settings.graphUserScopes);
  return response.token;
}

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

async function getTeamsAsync() {
  // Ensure client isn't undefined
  if (!_userClient) {
    throw new Error('Graph has not been initialized for user auth');
  }
  try {
  return _userClient.api('/me/joinedTeams')
    .select(['id', 'displayName', 'description'])
    .get();
}
 catch (e) { console.log(e)}
}

async function getChannelAsync(teamid) {
  // Ensure client isn't undefined
  if (!_userClient) {
    throw new Error('Graph has not been initialized for user auth');
  }
  try {
  const url = '/teams/' +teamid +'/channels'
  console.log(url);
  return _userClient.api(url)
    .select(['id', 'displayName', 'description'])
    .get();
}
 catch (e) { console.log(e)}
}

async function getFolderAsync(teamId,channelId) {
  // Ensure client isn't undefined
  if (!_userClient) {
    throw new Error('Graph has not been initialized for user auth');
  }
  try {
  const url = '/teams/' +teamId +'/channels/'+channelId+'/filesFolder'
  console.log(url);
  return _userClient.api(url)
   // .select(['id', 'name', 'webUrl'])
    .get();
}
 catch (e) { console.log(e)}
}
module.exports.getFolderAsync = getFolderAsync;
module.exports.getChannelAsync = getChannelAsync;
module.exports.getTeamsAsync = getTeamsAsync;
module.exports.getInboxAsync = getInboxAsync;
module.exports.getUserAsync = getUserAsync;
module.exports.getUserTokenAsync = getUserTokenAsync;
module.exports.initializeGraphForUserAuth = initializeGraphForUserAuth;


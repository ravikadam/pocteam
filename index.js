const readline = require('readline-sync');

const settings = require('./appSettings');
const graphHelper = require('./graphHelper');

async function main() {
  console.log('JavaScript Graph Tutorial');

  let choice = 0;
  let teamId,channelId

  // Initialize Graph
  initializeGraph(settings);

  // Greet the user by name
  await greetUserAsync();

  const choices = [
    'Display access token',
    'Show Teams',
    'Show Channels',
    'Get Folder Path',

  ];

  while (choice != -1) {
    choice = readline.keyInSelect(choices, 'Select an option', { cancel: 'Exit' });

    switch (choice) {
      case -1:
        // Exit
        console.log('Goodbye...');
        break;
      case 0:
        // Display access token
        await displayAccessTokenAsync();
        break;
      case 1:
        // List emails from user's inbox
       // await listInboxAsync();
        await listTeamAsync();

        break;
      case 2:
        // Send an email message
        await listChannelAsync();
        break;
      case 3:
        // List users
        await getFolderAsync();
        break;
      case 4:
        // Run any Graph code
        await makeGraphCallAsync();
        break;
      default:
        console.log('Invalid choice! Please try again.');
    }
  }
}

main();

function initializeGraph(settings) {
  graphHelper.initializeGraphForUserAuth(settings, (info) => {
    // Display the device code message to
    // the user. This tells them
    // where to go to sign in and provides the
    // code to use.
    console.log(info.message);
  });
}

async function greetUserAsync() {
  try {
    const user = await graphHelper.getUserAsync();
    console.log(`Hello, ${user.displayName}!`);
    // For Work/school accounts, email is in mail property
    // Personal accounts, email is in userPrincipalName
    console.log(`Email: ${user.mail}  ${user.userPrincipalName}  `);
  } catch (err) {
    console.log(`Error getting user: ${err}`);
  }
}

async function displayAccessTokenAsync() {
  try {
    const userToken = await graphHelper.getUserTokenAsync();
    console.log(`User token: ${userToken}`);
  } catch (err) {
    console.log(`Error getting user access token: ${err}`);
  }
}

async function listInboxAsync() {
  try {
    const messagePage = await graphHelper.getInboxAsync();
    const messages = messagePage.value;

    // Output each message's details
    for (const message of messages) {
      console.log(`Message: ${message.subject}`);
      console.log(`  From: ${message.from.emailAddress.name}`);
      console.log(`  Status: ${message.isRead ? 'Read' : 'Unread'}`);
      console.log(`  Received: ${message.receivedDateTime}`);
    }

    // If @odata.nextLink is not undefined, there are more messages
    // available on the server
    const moreAvailable = messagePage['@odata.nextLink'] != undefined;
    console.log(`\nMore messages available? ${moreAvailable}`);
  } catch (err) {
    console.log(`Error getting user's inbox: ${err}`);
  }
}

async function listTeamAsync() {
  try {
    const messagePage = await graphHelper.getTeamsAsync();
    const messages = messagePage.value;

    // Output each message's details
    for (const message of messages) {
      console.log(`TeamID: ${message.id}`);
      teamId = message.id;
      console.log(`  Name: ${message.displayName}`);
      console.log(`  Description: ${message.description}`);
    }
  } catch(e) {console.log(e)}
}

async function listChannelAsync() {
  try {
    const messagePage = await graphHelper.getChannelAsync(teamId);
    const messages = messagePage.value;

    // Output each message's details
    for (const message of messages) {
      console.log(`ChannelID: ${message.id}`);
      channelId = message.id;
      console.log(`  Name: ${message.displayName}`);
      console.log(`  Description: ${message.description}`);
    }
  } catch(e) {console.log(e)}
}

async function getFolderAsync() {
  try {
    const messagePage = await graphHelper.getFolderAsync(teamId,channelId);
    const messages = messagePage;
    // Output each message's details
      console.log(`Id: ${messages.id}`);
      console.log(`  Name: ${messages.name}`);
      console.log(`  WebURL: ${messages.webUrl}`);
    
  } catch(e) {console.log(e)}
}

async function sendMailAsync() {
  // TODO
}

async function listUsersAsync() {
  // TODO
}

async function makeGraphCallAsync() {
  // TODO
}

// Copyright (c) Microsoft Corporation. All rights reserved.
// Licensed under the MIT license.

const settings = {
  'clientId': 'f490b9d3-4f15-451a-b447-38567ea97c28',
  'clientSecret': 'YOUR_CLIENT_SECRET_HERE_IF_USING_APP_ONLY',
  'tenantId': 'YOUR_TENANT_ID_HERE_IF_USING_APP_ONLY',
  'authTenant': 'common',
  'graphUserScopes': [
    'user.read',
    'mail.read',
    'mail.send',
    'Team.ReadBasic.All',
    'Channel.ReadBasic.All',
    'Files.Read.All',
    'Directory.Read.All',
    'Group.Read.All'

  ]
};

module.exports = settings;


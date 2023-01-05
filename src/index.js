addEventListener("scheduled", (event) => {
  event.waitUntil(syncRiskyUserADGroups(event));
});

async function fetchAzureADToken() {

  if (!AZURE_AD_TENANT_ID) throw new Error('AZURE_AD_TENANT_ID environment variable is not defined')
  if (!AZURE_AD_CLIENT_ID) throw new Error('AZURE_AD_CLIENT_ID environment variable is not defined')
  if (!AZURE_AD_CLIENT_SECRET) throw new Error('AZURE_AD_CLIENT_SECRET environment variable is not defined')

  var tokenRequestURL = new URL('https://login.microsoftonline.com/' + AZURE_AD_TENANT_ID + '/oauth2/v2.0/token'); // https://learn.microsoft.com/en-us/graph/auth-v2-service

  tokenRequestBody = [];
  tokenRequestBody.push('client_id' + '=' + AZURE_AD_CLIENT_ID);
  tokenRequestBody.push('client_secret' + '=' + AZURE_AD_CLIENT_SECRET);
  tokenRequestBody.push('scope' + '=' + 'https://graph.microsoft.com/.default');
  tokenRequestBody.push('grant_type' + '=' + 'client_credentials');

  var tokenRequest = new Request(tokenRequestURL, {
    "method" : "POST",
    "headers" : {
      "Content-Type": "application/x-www-form-urlencoded"
    },
    "body" : tokenRequestBody.join('&')
  });

  var tokenResponse = await fetch(tokenRequest);
  if (tokenResponse.status !== 200) throw new Error('Failed to authenticate with Azure AD');
  var token = await tokenResponse.json();
  return token;
}

async function fetchAzureADGraph(method, url, body) {
  var request = new Request(url, {
    "headers" : new Headers(),
    "method": method,
    "body" : body
  })
  var token = await fetchAzureADToken();
  request.headers.set('Content-Type', 'application/json');
  request.headers.set('Authorization', 'Bearer ' + token.access_token);
  var response = await fetch(request);
  if (response.status == 403) throw new Error(await response.text())
  return response;
}

async function getRiskyUsers(riskLevel) {
  var riskyUsersReq = await fetchAzureADGraph('GET', "https://graph.microsoft.com/v1.0/identityProtection/riskyUsers?$filter=riskLevel eq '" + riskLevel + "'");
  var riskyUsers = await riskyUsersReq.json();
  var riskyUserIds = riskyUsers.value.map(function(user) {
    return user.id;
  });
  console.log('Found ' + (riskyUserIds.length) + ' ' + riskLevel + ' risk user(s).', riskyUserIds.join(','));
  return riskyUserIds;
}

async function getGroupByName(name) {
  var group = await fetchAzureADGraph('GET', "https://graph.microsoft.com/v1.0/groups?$filter=displayName eq '" + name + "'");
  return await group.json();
}

async function provisionGroup(name) {

  console.log('Provisioning risk group: ' + name);

  var existingGroup = await getGroupByName(name);
  if (existingGroup.value.length == 1) {
    console.log('Using existing risk group: ' + existingGroup.value[0].id);
    return existingGroup.value[0];
  } else {

    console.log('Risk group does not exist and will be created...');
    
    var newGroupRequest = await fetchAzureADGraph('POST', 'https://graph.microsoft.com/v1.0/groups', JSON.stringify({
      "displayName" : name,
      "description" : "Group synchronized with Azure Risky Users API",
      "mailEnabled" : false,
      "mailNickname" : name,
      "groupTypes": [],
      "securityEnabled" : true
    }));

    var newGroup = await newGroupRequest.json();
    console.log('Created risk group: ' + newGroup.id);
    return newGroup;
  }
}

async function getGroupMembers(id) {
  console.log('Fetching group members: ' + id)
  var groupMembersReq = await fetchAzureADGraph('GET', 'https://graph.microsoft.com/v1.0/groups/' + id + '/members');
  var groupMembers = await groupMembersReq.json();
  var groupMemberIds = groupMembers.value.map(function(user) {
    return user.id;
  });
  console.log('Found ' + groupMemberIds.length + ' member(s).', groupMemberIds.join(','));
  return groupMemberIds;
}

async function addMembers(groupId, userIds) {
  var members = userIds.map(function(user) {
    return "https://graph.microsoft.com/v1.0/directoryObjects/" + user
  })

  console.log('Adding members to group ' + groupId);
  var addMembersReq = await fetchAzureADGraph('PATCH', 'https://graph.microsoft.com/v1.0/groups/' + groupId, JSON.stringify({"members@odata.bind": members}));
  if (addMembersReq.status == 204) {
    console.log('Successfully added user(s) to group.', userIds.join(','))
  } else {
    console.error('An error occurred while adding users to group.', await addMembersReq.json());
  }

}

async function removeMembers(groupId, userIds) {
  var members = userIds.map(function(user) {
    return "https://graph.microsoft.com/v1.0/directoryObjects/" + user
  })
  var removeMembersReq = await fetchAzureADGraph('DELETE', 'https://graph.microsoft.com/v1.0/groups/' + groupId, JSON.stringify({"members@odata.bind": members}));
  if (removeMembersReq.status == 204) {
    console.log('Successfully removed user(s) from group.', userIds.join(','))
  } else {
    console.error('An error occurred while removing users to group.', await removeMembersReq.json());
  }
}

async function syncGroup(riskLevel) {

  console.log('Started synchronizing ' + riskLevel + ' risky users.')
  var riskyUsers = await getRiskyUsers(riskLevel);
  var group = await provisionGroup('IdentityProtection-RiskyUser-RiskLevel-' + riskLevel);
  var groupMembers = await getGroupMembers(group.id);

  var staleMembers = groupMembers.filter(function(user) {
    return riskyUsers.indexOf(user) === -1;
  });
  
  var newMembers = riskyUsers.filter(function(user) {
    return groupMembers.indexOf(user) === -1;
  });

  if (newMembers.length > 0) {
    await addMembers(group.id, newMembers);
  } else {
    console.log('No new users to add to group.')
  }

  if (staleMembers.length > 0) {
    await removeMembers(group.id, staleMembers);
  } else {
    console.log('No stale users to remove from group.')
  }

  console.log('Finished synchronizing ' + riskLevel + 'users.')

}

async function syncRiskyUserADGroups(event) {
  
  console.log('Started Azure AD Risky User to Group Synchronization.')
  await syncGroup('high')
  await syncGroup('medium')
  await syncGroup('low')
  console.log('Finished Azure AD Risky User to Group Synchronization.')

}

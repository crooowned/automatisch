import { URLSearchParams } from 'url';

const verifyCredentials = async ($) => {
  const oauthRedirectUrlField = $.app.auth.fields.find(
    (field) => field.key === 'oAuthRedirectUrl'
  );

  if(!$.auth.data.code) {
    throw new Error('No code found');
  }

  // Were getting first refreshtoken here
  const params = new URLSearchParams({
    client_id: $.auth.data.clientId,
    redirect_uri: oauthRedirectUrlField.value,
    scope: $.app.scopes.join(' '),
    client_secret: $.auth.data.clientSecret,
    code: $.auth.data.code,
    grant_type: 'authorization_code',
  });

  const { data: tokenData } = await $.http.post(
    'https://login.microsoftonline.com/common/oauth2/v2.0/token',
    params.toString(),
    {
      additionalProperties: {
        skipAddingBaseUrl: true,
      },
    }
  );

  // Get user info to set screen name
  const { data: userData } = await $.http.get('https://graph.microsoft.com/v1.0/me', {
    headers: {
      Authorization: `Bearer ${tokenData.access_token}`,
    },
    additionalProperties: {
      skipAddingBaseUrl: true,
    },
  });
  
  await $.auth.set({
    accessToken: tokenData.access_token,
    refreshToken: tokenData.refresh_token,
    tokenType: tokenData.token_type,
    expiresIn: tokenData.expires_in,
    scope: tokenData.scope,
    screenName: userData.userPrincipalName || userData.displayName,
  });
};

export default verifyCredentials;

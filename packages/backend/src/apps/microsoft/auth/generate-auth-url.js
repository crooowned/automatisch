import { URLSearchParams } from 'url';


export default async function generateAuthUrl($) {
  const oauthRedirectUrlField = $.app.auth.fields.find(
    (field) => field.key === 'oAuthRedirectUrl'
  );
  
  const params = new URLSearchParams({
    client_id: $.auth.data.clientId,
    redirect_uri: oauthRedirectUrlField.value,
    response_type: 'code',
    scope: $.app.scopes.join(' '),
    response_mode: 'query',
  });

  // Store client ID and secret for later use in verification
  await $.auth.set({
    clientId: $.auth.data.clientId,
    clientSecret: $.auth.data.clientSecret,
  });
  await $.auth.set({ url: `https://login.microsoftonline.com/common/oauth2/v2.0/authorize?${params.toString()}` });

  return true;  // Indicate that the URL was generated successfully
}

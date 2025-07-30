export default async function refreshToken($) {
  if (!$.auth.data.refreshToken) {
    console.warn("No refresh token available, cannot refresh.");
    return false;
  }

  const params = {
    client_id: $.auth.data.clientId,
    client_secret: $.auth.data.clientSecret,
    grant_type: 'refresh_token',
    refresh_token: $.auth.data.refreshToken,
  };

  const { data } = await $.http.post(
    'https://login.microsoftonline.com/common/oauth2/v2.0/token',
    new URLSearchParams(params),
    {
      headers: {
        'Content-Type': 'application/x-www-form-urlencoded',
      },
    }
  );

  await $.auth.set({
    accessToken: data.access_token,
    refreshToken: data.refresh_token,
    expiresIn: data.expires_in,
  });

  return true;
}

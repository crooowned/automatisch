export default async function addAuthHeader($, requestConfig) {
  if ($.auth.data.accessToken) {
    requestConfig.headers.Authorization = `Bearer ${$.auth.data.accessToken}`;
  }

  return requestConfig;
}

import generateAuthUrl from './generate-auth-url.js';
import verifyCredentials from './verify-credentials.js';
import isStillVerified from './is-still-verified.js';
import refreshToken from './refresh-token.js';

export default {
  fields: [
    {
      key: 'oAuthRedirectUrl',
      label: 'OAuth Redirect URL',
      type: 'string',
      required: true,
      readOnly: true,
      value: '{WEB_APP_URL}/app/microsoft/connections/add',
      placeholder: null,
      description: 'When asked to input an OAuth callback or redirect URL in Microsoft Azure App Registration, enter the URL above.',
      docUrl: '{DOCS_URL}/apps/microsoft/connection#oauth-redirect-url',
      clickToCopy: true,
    },
    {
      key: 'clientId',
      label: 'Client ID',
      type: 'string',
      required: true,
      readOnly: false,
      value: null,
      placeholder: null,
      description: 'The application (client) ID from your Azure App Registration',
      docUrl: '{DOCS_URL}/apps/microsoft/connection#client-id',
      clickToCopy: false,
    },
    {
      key: 'clientSecret',
      label: 'Client Secret',
      type: 'string',
      required: true,
      readOnly: false,
      value: null,
      placeholder: null,
      description: 'The client secret value from your Azure App Registration',
      docUrl: '{DOCS_URL}/apps/microsoft/connection#client-secret',
      clickToCopy: false,
    }
  ],

  generateAuthUrl,
  verifyCredentials,
  isStillVerified,
  refreshToken,
};

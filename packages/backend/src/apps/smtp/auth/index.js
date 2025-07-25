import verifyCredentials from './verify-credentials.js';
import isStillVerified from './is-still-verified.js';

export default {
  fields: [
    {
      key: 'screenName',
      label: 'Screen Name',
      type: 'string',
      required: true,
      readOnly: false,
      value: null,
      placeholder: null,
      description: 'Screen name of your connection to be shown in the UI.',
      clickToCopy: false,
    },
    {
      key: 'host',
      label: 'Host',
      type: 'string',
      required: true,
      readOnly: false,
      value: null,
      placeholder: null,
      description: 'The host information Automatisch will connect to.',
      docUrl: 'https://automatisch.io/docs/smtp#host',
      clickToCopy: false,
    },
    {
      key: 'username',
      label: 'Email/Username',
      type: 'string',
      required: false,
      readOnly: false,
      value: null,
      placeholder: null,
      description: 'Your SMTP login credentials.',
      docUrl: 'https://automatisch.io/docs/smtp#username',
      clickToCopy: false,
    },
    {
      key: 'password',
      label: 'Password',
      type: 'string',
      required: false,
      readOnly: false,
      value: null,
      placeholder: null,
      description: null,
      docUrl: 'https://automatisch.io/docs/smtp#password',
      clickToCopy: false,
    },
    {
      key: 'useTls',
      label: 'Use TLS?',
      type: 'dropdown',
      required: false,
      readOnly: false,
      value: false,
      placeholder: null,
      description: null,
      docUrl: 'https://automatisch.io/docs/smtp#use-tls',
      clickToCopy: false,
      options: [
        {
          label: 'Yes',
          value: true,
        },
        {
          label: 'No',
          value: false,
        },
      ],
    },
    {
      key: 'port',
      label: 'Port',
      type: 'string',
      required: false,
      readOnly: false,
      value: '25',
      placeholder: null,
      description: null,
      docUrl: 'https://automatisch.io/docs/smtp#port',
      clickToCopy: false,
    },
  ],
  verifyCredentials,
  isStillVerified,
};

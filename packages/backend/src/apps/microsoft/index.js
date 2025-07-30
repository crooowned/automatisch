import defineApp from "../../helpers/define-app.js";
import auth from './auth/index.js';
import triggers from './triggers/index.js';
import addAuthHeader from './common/add-auth-header.js';
import setBaseUrl from './common/set-base-url.js';
import dynamicData from './dynamic-data/index.js';
import actions from './actions/index.js';
import dynamicFields from "./dynamic-fields/index.js";

export default defineApp({
    name: 'Microsoft',
    key: 'microsoft',
    iconUrl: '{BASE_URL}/apps/microsoft/assets/favicon.svg',
    authDocUrl: 'https://anysite.me/',
    baseUrl: 'https://login.microsoftonline.com',
    apiBaseUrl: 'https://graph.microsoft.com/v1.0',
    scopes: [
        'email',
        'Files.ReadWrite',
        'Files.ReadWrite.All',
        'Mail.Read',
        'Mail.Read.Shared',
        'Mail.ReadBasic.Shared',
        'Mail.ReadWrite',
        'Mail.ReadWrite.Shared',
        'Mail.Send',
        'Mail.Send.Shared',
        'MailboxFolder.Read',
        'MailboxFolder.ReadWrite',
        'offline_access',
        'openid',
        'profile',
        'User-Mail.ReadWrite.All',
        'User.Read',
        'User.Read.All'
    ],
    primaryColor: '#0078d4',
    supportsConnections: true,
    beforeRequest: [setBaseUrl, addAuthHeader],
    auth,
    triggers,
    dynamicData,
    dynamicFields,
    actions
});
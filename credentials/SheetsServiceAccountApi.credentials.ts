import type {
    ICredentialTestRequest,
    ICredentialType,
    INodeProperties,
} from 'n8n-workflow';

export class SheetsServiceAccountApi implements ICredentialType {
    name = 'sheetsServiceAccountApi';
    displayName = 'Sheets Service Account';
    documentationUrl = 'https://github.com/ariadng/sheets';

    properties: INodeProperties[] = [
        {
            displayName: 'OAuth Client ID',
            name: 'oauthClientId',
            type: 'string',
            default: '',
            required: true,
            description: 'OAuth 2.0 Client ID from Google Cloud Console. Create a Desktop app credential.',
        },
        {
            displayName: 'OAuth Client Secret',
            name: 'oauthClientSecret',
            type: 'string',
            typeOptions: {
                password: true,
            },
            default: '',
            required: true,
            description: 'OAuth 2.0 Client Secret from Google Cloud Console.',
        },
        {
            displayName: 'Service Account JSON',
            name: 'serviceAccountJson',
            type: 'string',
            typeOptions: {
                password: true,
                rows: 10,
            },
            default: '',
            required: true,
            description: 'Paste the entire JSON key file content from Google Cloud Console. The JSON should contain client_email and private_key fields.',
            placeholder: '{"type": "service_account", "project_id": "...", ...}',
        },
    ];

    // Note: We don't use authenticate here because the @ariadng/sheets library
    // handles authentication internally. The credential is just storage.

    // Credential test - validates the JSON structure
    test: ICredentialTestRequest = {
        request: {
            // This is a dummy request - actual validation happens in the node
            // because @ariadng/sheets handles the JWT creation internally
            baseURL: 'https://sheets.googleapis.com',
            url: '/v4/spreadsheets/test-credential-validation',
            skipSslCertificateValidation: false,
        },
    };
}

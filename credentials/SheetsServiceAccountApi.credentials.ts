import type {
    ICredentialType,
    INodeProperties,
} from 'n8n-workflow';

export class SheetsServiceAccountApi implements ICredentialType {
    name = 'sheetsServiceAccountApi';
    displayName = 'Sheets Service Account';
    documentationUrl = 'https://github.com/ariadng/sheets';

    properties: INodeProperties[] = [
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

    // Note: Credential validation happens when the node executes.
    // The @ariadng/sheets library handles JWT auth internally.
}

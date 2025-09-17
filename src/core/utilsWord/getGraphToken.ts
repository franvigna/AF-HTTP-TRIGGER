// src/core/utilsWord/getGraphToken.ts
export async function getGraphToken(): Promise<string | null> {
    const clientId = process.env.GRAPH_CLIENT_ID!;
    const clientSecret = process.env.GRAPH_CLIENT_SECRET!;
    const tenantId = process.env.GRAPH_TENANT_ID!;

    const url = `https://login.microsoftonline.com/${tenantId}/oauth2/v2.0/token`;

    const params = new URLSearchParams();
    params.append('client_id', clientId);
    params.append('client_secret', clientSecret);
    params.append('grant_type', 'client_credentials');
    params.append('scope', 'https://graph.microsoft.com/.default');

    try {
        const response = await fetch(url, {
            method: 'POST',
            headers: { 'Content-Type': 'application/x-www-form-urlencoded' },
            body: params.toString(),
        });

        const data = await response.json();

        if (!response.ok) {
            console.error('❌ Error al obtener token:', data);
            return null;
        }

        return data.access_token ?? null;
    } catch (err: any) {
        console.error('❌ Error de red o código:', err.message || err);
        return null;
    }
}

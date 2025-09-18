// src/core/utilsWord/getSharePointToken.ts
export async function getSharePointToken(
    siteHostname: string
): Promise<string | null> {
    const clientId = process.env.SPO_CLIENT_ID!;
    const clientSecret = process.env.SPO_CLIENT_SECRET!;
    const tenantId = process.env.SPO_TENANT_ID!;
    // audiencia para SharePoint REST en tu tenant:
    const scope = `https://${siteHostname}/.default`;

    const url = `https://login.microsoftonline.com/${tenantId}/oauth2/v2.0/token`;
    const params = new URLSearchParams();
    params.append('client_id', clientId);
    params.append('client_secret', clientSecret);
    params.append('grant_type', 'client_credentials');
    params.append('scope', scope);

    try {
        const response = await fetch(url, {
            method: 'POST',
            body: params as any,
        });
        const data = await response.json();
        if (!response.ok) {
            console.error('❌ Error al obtener token SPO:', data);
            return null;
        }
        return data.access_token ?? null;
    } catch (err: any) {
        console.error('❌ Error de red o código SPO:', err?.message || err);
        return null;
    }
}

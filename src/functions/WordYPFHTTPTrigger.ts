import {
    app,
    HttpRequest,
    HttpResponseInit,
    InvocationContext,
} from '@azure/functions';
import { Packer } from 'docx';
import { buildDoc } from '../core/utilsWord/buildDoc';
import { GraphListService } from '../core/datosSharepoint/listaConGraph';
import type { BuildDocInput } from '../core/utilsWord/types';
import { getGraphToken } from '../core/utilsWord/getGraphToken';

const SITE_HOSTNAME = 'circo.sharepoint.com';
const SITE_ID = 'b07b24d5-c53f-402e-8756-c4bb17dd9997';
const WEB_ID = '6df63e08-1c14-4aaf-9b7a-b1cff7722134';
const LIST_NAME = 'Word';

export async function WordYPFHTTPTrigger(
    request: HttpRequest,
    context: InvocationContext
): Promise<HttpResponseInit> {
    context.log('Generating Word with SharePoint data...');

    try {
        // ✅ 1. Obtener el token dinámicamente usando client credentials
        const token = await getGraphToken();
        if (!token) {
            context.log(
                'ERROR: No se pudo obtener el token de Microsoft Graph.'
            );
            return { status: 500, body: 'Error al obtener token de Graph.' };
        }

        // ✅ 2. Crear servicio con token
        const service = new GraphListService(
            token,
            SITE_HOSTNAME,
            SITE_ID,
            WEB_ID
        );
        const novedadesData = await service.fetchNovedades(LIST_NAME);

        // ✅ 3. Armar input para buildDoc
        const docInput: BuildDocInput = {
            confidentialityLabel: 'Interno YPF',
            novedades: novedadesData.novedad.map((n) => ({
                areaNovedad: n.areaNovedad,
                ...n.items[0],
                sectorGeneral: novedadesData.area,
            })),
        };

        const doc = await buildDoc(docInput);
        const buffer = await Packer.toBuffer(doc);

        // ✅ 4. Devolver el Word como descarga
        return {
            status: 200,
            headers: {
                'Content-Type':
                    'application/vnd.openxmlformats-officedocument.wordprocessingml.document',
                'Content-Disposition': 'attachment; filename=YPF-Reporte.docx',
            },
            body: buffer,
        };
    } catch (error: any) {
        context.log('ERROR:', error.message || error);
        return {
            status: 500,
            body: `Error interno al generar el documento: ${
                error.message || error
            }`,
        };
    }
}

app.http('WordYPFHTTPTrigger', {
    methods: ['GET'],
    authLevel: 'anonymous',
    handler: WordYPFHTTPTrigger,
});

import {
    app,
    HttpRequest,
    HttpResponseInit,
    InvocationContext,
} from '@azure/functions';

import { Packer } from 'docx';
import { buildDoc } from '../core/utilsWord/buildDoc';
import type { BuildDocInput } from '../core/utilsWord/types';

// 🚀 Simulamos la obtención de token (no se usa en mock, solo de ejemplo)
async function getMockGraphToken(): Promise<string> {
    return 'mocked_graph_token';
}

// 🚀 Simulamos el servicio de SharePoint con estructura compatible con `buildDoc`
async function fetchMockSharePointData(): Promise<BuildDocInput> {
    return {
        confidentialityLabel: 'Interno YPF',
        novedades: [
            {
                sectorGeneral: 'Refinería',
                areaNovedad: 'Seguridad',
                tituloNovedad: 'Cambio en los procedimientos de evacuación',
                detalleNovedad:
                    'Se han actualizado los procedimientos según las nuevas normativas ISO.',
                resumenHtml:
                    '<p><strong>Actualización:</strong> Se reemplazaron las salidas de emergencia del sector 3.</p>',
                imagenesNovedad: [
                    'https://upload.wikimedia.org/wikipedia/commons/thumb/4/47/PNG_transparency_demonstration_1.png/640px-PNG_transparency_demonstration_1.png',
                ],
            },
            {
                sectorGeneral: 'Refinería',
                areaNovedad: 'Mantenimiento',
                tituloNovedad: 'Nueva maquinaria instalada',
                detalleNovedad:
                    'Se completó la instalación del nuevo compresor Siemens modelo X300.',
                resumenHtml:
                    '<ul><li>Mayor eficiencia</li><li>Menor consumo energético</li></ul>',
                imagenesNovedad: [],
            },
        ],
    };
}

// ✅ Azure Function con estructura real
export async function GraphMockFunction(
    request: HttpRequest,
    context: InvocationContext
): Promise<HttpResponseInit> {
    context.log('📄 Generando Word con MOCK de SharePoint...');

    try {
        // Simulación del token (no se usa pero se respeta la estructura del real)
        const token = await getMockGraphToken();
        context.log('✅ Token mock obtenido:', token);

        // Simulación de obtención de datos desde SharePoint
        const inputData = await fetchMockSharePointData();

        // Generamos el Word con la misma función que usarías en producción
        const doc = await buildDoc(inputData);
        const buffer = await Packer.toBuffer(doc);

        return {
            status: 200,
            headers: {
                'Content-Type':
                    'application/vnd.openxmlformats-officedocument.wordprocessingml.document',
                'Content-Disposition':
                    'attachment; filename=YPF-Reporte-Mock.docx',
            },
            body: buffer,
        };
    } catch (error: any) {
        context.log('❌ ERROR:', error.message || error);
        return {
            status: 500,
            body: `Error inesperado: ${error.message || error}`,
        };
    }
}

app.http('GraphMockFunction', {
    methods: ['GET'],
    authLevel: 'anonymous',
    handler: GraphMockFunction,
});

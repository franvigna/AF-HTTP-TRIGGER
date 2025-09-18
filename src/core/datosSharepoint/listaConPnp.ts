// src/core/datosSharepoint/listaConPnp.ts
import { SPFI, spfi } from '@pnp/sp';
import { BearerToken } from '@pnp/queryable';

// Extensiones de PnP que vamos a usar
import '@pnp/sp/webs';
import '@pnp/sp/lists';
import '@pnp/sp/items';
import '@pnp/sp/attachments';

export type ListaPruebaFields = {
    SectorGeneral?: string;
    AreaNovedad?: string;
    TituloNovedad?: string;
    DetalleNovedad?: string;
    Resumen?: string;
    Title?: string;
};

export type NovedadItem = {
    sectorGeneral: string;
    tituloNovedad: string;
    resumenHtml: string;
    detalleNovedad: string;
    imagenesNovedad: string[];
};

export type NovedadesData = {
    area: string;
    novedad: Array<{ areaNovedad: string; items: NovedadItem[] }>;
};

export class PnpListService {
    private sp: SPFI;
    private siteOrigin: string;

    /**
     * @param accessToken Token AAD para SharePoint (audiencia https://{tenant}.sharepoint.com/.default)
     * @param siteUrl     URL del sitio, ej: "https://circo.sharepoint.com/sites/NombreDelSitio"
     */
    constructor(private accessToken: string, private siteUrl: string) {
        this.sp = spfi(siteUrl).using(BearerToken(accessToken));
        this.siteOrigin = new URL(siteUrl).origin;
    }

    /** Equivalente al fetchNovedades con Graph, pero via REST de SharePoint usando PnPjs */
    public async fetchNovedades(listTitle: string): Promise<NovedadesData> {
        const list = this.sp.web.lists.getByTitle(listTitle);

        // Verificamos que la lista exista (si no, lanzará)
        await list.select('Id')();

        // Traemos hasta 200 items con los campos necesarios (+ flag Attachments)
        const rawItems = await list.items
            .select(
                'Id,Title,SectorGeneral,AreaNovedad,TituloNovedad,DetalleNovedad,Resumen,Attachments'
            )
            .top(200)();

        // Procesamos cada item y resolvemos adjuntos en paralelo
        const processed = await Promise.all(
            rawItems.map(async (it: any) => {
                const f: ListaPruebaFields = {
                    SectorGeneral: it.SectorGeneral,
                    AreaNovedad: it.AreaNovedad,
                    TituloNovedad: it.TituloNovedad,
                    DetalleNovedad: it.DetalleNovedad,
                    Resumen: it.Resumen,
                    Title: it.Title,
                };

                const areaNovedad = (f.AreaNovedad ?? 'General').trim();
                const titulo = (
                    f.TituloNovedad ??
                    f.Title ??
                    '(Sin título)'
                ).trim();
                const detalleHtml = f.DetalleNovedad ?? '';
                const resumen = f.Resumen ?? '';
                const sectorGeneral = (
                    f.SectorGeneral ??
                    f.Title ??
                    'General'
                ).trim();

                // Adjuntos como imágenes (equivalente a /driveItem/children)
                const urls: string[] = [];
                if (it.Attachments === true) {
                    const atts = await list.items
                        .getById(it.Id)
                        .attachmentFiles();
                    for (const a of atts as any[]) {
                        // Filtramos por extensión de imagen
                        if (
                            /\.(png|jpe?g|gif|bmp|webp|tiff?)$/i.test(
                                a.FileName
                            )
                        ) {
                            // ServerRelativeUrl -> URL absoluta
                            urls.push(
                                `${this.siteOrigin}${a.ServerRelativeUrl}`
                            );
                        }
                    }
                }

                const novedad: NovedadItem = {
                    tituloNovedad: titulo,
                    detalleNovedad: detalleHtml,
                    resumenHtml: resumen,
                    imagenesNovedad: Array.from(new Set(urls)),
                    sectorGeneral,
                };

                return { areaNovedad, novedad, sectorGeneral };
            })
        );

        // Agrupación por AreaNovedad y determinación de SectorGeneral más frecuente
        const groups = new Map<string, NovedadItem[]>();
        const sectorVals: string[] = [];

        for (const row of processed) {
            sectorVals.push(row.sectorGeneral);
            const arr = groups.get(row.areaNovedad) ?? [];
            arr.push(row.novedad);
            groups.set(row.areaNovedad, arr);
        }

        const area = this.pickMostFrequent(sectorVals);

        const novedad = Array.from(groups.entries())
            .map(([areaNovedad, items]) => ({ areaNovedad, items }))
            .sort((a, b) => a.areaNovedad.localeCompare(b.areaNovedad));

        return { area, novedad };
    }

    private pickMostFrequent(values: string[]): string {
        if (!values.length) return 'General';
        const counts: Record<string, number> = {};
        for (const v of values) counts[v] = (counts[v] || 0) + 1;

        let best = 'General';
        let max = -1;
        for (const [k, v] of Object.entries(counts)) {
            if (v > max) {
                max = v;
                best = k;
            }
        }
        return best;
    }
}

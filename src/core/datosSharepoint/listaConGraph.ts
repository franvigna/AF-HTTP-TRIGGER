import { Client } from '@microsoft/microsoft-graph-client';
import 'isomorphic-fetch'; // required for Graph SDK in Node.js

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

type GraphListItem = {
    id: string;
    fields: ListaPruebaFields;
};

export class GraphListService {
    private client: Client;

    constructor(
        private accessToken: string,
        private siteHostname: string, // e.g., tenant.sharepoint.com
        private siteId: string, // site ID (not URL)
        private webId: string // web ID (needed for building site identifier)
    ) {
        this.client = Client.init({
            authProvider: (done) => done(null, accessToken),
        });
    }

    public async fetchNovedades(
        listDisplayName: string
    ): Promise<NovedadesData> {
        const siteUniqueId = `${this.siteHostname},${this.siteId},${this.webId}`;
        const list = await this.getList(siteUniqueId, listDisplayName);
        const items = await this.getItems(siteUniqueId, list?.id);

        const groups = new Map<string, NovedadItem[]>();

        for (const it of items as GraphListItem[]) {
            const f: ListaPruebaFields = it.fields || {};
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

            const urls: string[] = [];

            if (list?.id) {
                try {
                    const resp = await this.client
                        .api(
                            `/sites/${siteUniqueId}/lists/${list.id}/items/${it.id}/driveItem/children`
                        )
                        .select('id,name,@microsoft.graph.downloadUrl,file')
                        .get();

                    const files = resp?.value ?? [];
                    for (const ch of files) {
                        const mime = ch?.file?.mimeType || '';
                        const dl = ch?.['@microsoft.graph.downloadUrl'];
                        if (dl && /^image\//i.test(mime)) urls.push(dl);
                    }
                } catch {
                    // ignorar
                }
            }

            const arr = groups.get(areaNovedad) ?? [];
            arr.push({
                tituloNovedad: titulo,
                detalleNovedad: detalleHtml,
                resumenHtml: resumen,
                imagenesNovedad: Array.from(new Set(urls)),
                sectorGeneral,
            });
            groups.set(areaNovedad, arr);
        }

        const sectorGeneral = this.pickMostFrequent(
            items.map((it: GraphListItem) =>
                (
                    it.fields?.SectorGeneral ??
                    it.fields?.Title ??
                    'General'
                ).trim()
            )
        );

        const novedad = Array.from(groups.entries())
            .map(([areaNovedad, items]) => ({ areaNovedad, items }))
            .sort((a, b) => a.areaNovedad.localeCompare(b.areaNovedad));

        return { area: sectorGeneral, novedad };
    }

    private async getList(siteUniqueId: string, displayName: string) {
        const response = await this.client
            .api(`/sites/${siteUniqueId}/lists`)
            .select('id,name,displayName')
            .get();

        const lists: any[] = response?.value ?? [];

        const match = lists.find(
            (l: any) =>
                (l.displayName || '').toLowerCase() ===
                    displayName.toLowerCase() ||
                (l.name || '').toLowerCase() === displayName.toLowerCase()
        );

        return match
            ? { id: match.id, name: match.displayName || match.name || '' }
            : null;
    }

    private async getItems(siteUniqueId: string, listId?: string) {
        if (!listId) return [];

        const response = await this.client
            .api(`/sites/${siteUniqueId}/lists/${listId}/items`)
            .expand(
                'fields($select=Title,SectorGeneral,AreaNovedad,TituloNovedad,DetalleNovedad,Resumen)'
            )
            .top(200)
            .get();

        return response?.value ?? [];
    }

    private pickMostFrequent(values: string[]): string {
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

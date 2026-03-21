import { IInputs, IOutputs } from "./generated/ManifestTypes";

interface KpiRecord {
    id: string;
    name: string;
    status: number;
    failuretime: Date | null;
    warningtime: Date | null;
    succeededon: Date | null;
    applicablefrom: Date | null;
    pausedon: Date | null;
    terminalstatereached: boolean;
}

const SLA_STATUS = {
    IN_PROGRESS: 0,
    NONCOMPLIANT: 1,
    NEARING: 2,
    PAUSED: 3,
    SUCCEEDED: 4,
    CANCELLED: 5,
};

const R_OUTER = 20;
const R_INNER = 16;
const CIRC_OUTER = 2 * Math.PI * R_OUTER;
const CIRC_INNER = 2 * Math.PI * R_INNER;

export class ModernSlaTimerControl implements ComponentFramework.StandardControl<IInputs, IOutputs> {
    private _container!: HTMLDivElement;
    private _tickInterval: number | null = null;
    private _kpis: KpiRecord[] = [];
    private _enableNegativeTimer: boolean = false;
    private _enriched: boolean = false;

    public init(
        context: ComponentFramework.Context<IInputs>,
        _notifyOutputChanged: () => void,
        _state: ComponentFramework.Dictionary,
        container: HTMLDivElement
    ): void {
        this._container = container;
        this._container.classList.add("msla-container");
        this._enableNegativeTimer = context.parameters.EnableNegativeTimer?.raw === "1";
        this._tickInterval = window.setInterval(() => this._tick(), 1000);
    }

    public updateView(context: ComponentFramework.Context<IInputs>): void {
        const ds = context.parameters.dataSetGrid_1;
        if (!ds || ds.loading) return;

        this._enableNegativeTimer = context.parameters.EnableNegativeTimer?.raw === "1";
        this._kpis = this._extractKpis(ds);
        this._renderAll();

        if (!this._enriched && this._kpis.length > 0) {
            this._enriched = true;
            this._enrichFromApi();
        }
    }

    public getOutputs(): IOutputs {
        return {};
    }

    public destroy(): void {
        if (this._tickInterval !== null) window.clearInterval(this._tickInterval);
    }

    // ── Data extraction ──────────────────────────────────────

    private _extractKpis(ds: ComponentFramework.PropertyTypes.DataSet): KpiRecord[] {
        const kpis: KpiRecord[] = [];
        for (const id of ds.sortedRecordIds) {
            const rec = ds.records[id];
            const prev = this._kpis.find((k) => k.id === id);
            kpis.push({
                id,
                name: this._str(rec, "name"),
                status: this._num(rec, "status"),
                failuretime: this._date(rec, "failuretime"),
                warningtime: this._date(rec, "warningtime"),
                succeededon: this._date(rec, "succeededon"),
                applicablefrom:
                    prev?.applicablefrom ??
                    this._date(rec, "applicablefromvalue") ??
                    this._date(rec, "applicablefrom") ??
                    this._date(rec, "createdon"),
                pausedon: this._date(rec, "pausedon"),
                terminalstatereached:
                    this._str(rec, "terminalstatereached") === "true" ||
                    this._str(rec, "terminalstatereached") === "1",
            });
        }
        return kpis;
    }

    private _enrichFromApi(): void {
        const ids = this._kpis.map((k) => k.id.replace(/[{}]/g, ""));
        const filter = ids.map((i) => `slakpiinstanceid eq ${i}`).join(" or ");
        const url = `/api/data/v9.2/slakpiinstances?$select=slakpiinstanceid,applicablefromvalue,createdon&$filter=${filter}`;

        fetch(url, {
            credentials: "same-origin",
            headers: {
                Accept: "application/json",
                "OData-MaxVersion": "4.0",
                "OData-Version": "4.0",
            },
        })
            .then((r) => {
                if (r.ok) return r.json();
                throw new Error("fetch failed");
            })
            .then((data) => {
                for (const e of data.value) {
                    const eid: string = (e["slakpiinstanceid"] ?? "").toLowerCase();
                    const kpi = this._kpis.find(
                        (k) => k.id.replace(/[{}]/g, "").toLowerCase() === eid
                    );
                    if (kpi) {
                        const d = this._toDate(e["applicablefromvalue"]) ?? this._toDate(e["createdon"]);
                        if (d) kpi.applicablefrom = d;
                    }
                }
                this._renderAll();
                return data;
            })
            .catch(() => {});
    }

    private _toDate(v: unknown): Date | null {
        if (!v) return null;
        const d = v instanceof Date ? v : new Date(String(v));
        return isNaN(d.getTime()) ? null : d;
    }

    // ── Record helpers ───────────────────────────────────────

    private _str(rec: ComponentFramework.PropertyHelper.DataSetApi.EntityRecord, col: string): string {
        try {
            return rec.getFormattedValue(col) ?? rec.getValue(col)?.toString() ?? "";
        } catch {
            return "";
        }
    }

    private _num(rec: ComponentFramework.PropertyHelper.DataSetApi.EntityRecord, col: string): number {
        try {
            const v = rec.getValue(col);
            return typeof v === "number" ? v : parseInt(v?.toString() ?? "0", 10);
        } catch {
            return 0;
        }
    }

    private _date(rec: ComponentFramework.PropertyHelper.DataSetApi.EntityRecord, col: string): Date | null {
        try {
            const v = rec.getValue(col);
            if (!v) return null;
            const d = v instanceof Date ? v : new Date(v.toString());
            return isNaN(d.getTime()) ? null : d;
        } catch {
            return null;
        }
    }

    // ── Rendering ────────────────────────────────────────────

    private _renderAll(): void {
        if (!this._container) return;
        if (this._kpis.length === 0) {
            this._container.innerHTML = `<div class="msla-empty">No SLA timers</div>`;
            return;
        }
        this._container.innerHTML = this._kpis.map((k, i) => this._renderCard(k, i)).join("");
    }

    private _renderCard(kpi: KpiRecord, idx: number): string {
        const sk = this._statusKey(kpi);
        const visual = this._renderVisual(kpi, idx, sk);
        const countdown = this._calcCountdown(kpi);
        const badge = this._badgeText(sk);
        return `<div class="msla-card" data-status="${sk}"><div class="msla-visual">${visual}</div><div class="msla-info"><span class="msla-kpi-name">${this._escapeHtml(kpi.name)}</span><div class="msla-card-row"><span class="msla-countdown" data-status="${sk}" data-idx="${idx}">${countdown}</span><span class="msla-badge" data-status="${sk}"><span class="msla-badge-dot"></span>${badge}</span></div></div></div>`;
    }

    private _renderVisual(kpi: KpiRecord, idx: number, sk: string): string {
        switch (sk) {
            case "succeeded":
                return '<svg viewBox="0 0 44 44" class="msla-svg-icon msla-glow-green"><circle cx="22" cy="22" r="18" fill="#107c10"/><path d="M14 22.5l5.5 5.5 10-11" fill="none" stroke="#fff" stroke-width="2.8" stroke-linecap="round" stroke-linejoin="round" class="msla-check-draw"/></svg>';
            case "noncompliant":
                return '<svg viewBox="0 0 44 44" class="msla-svg-icon msla-glow-red"><circle cx="22" cy="22" r="18" fill="#d13438"/><path d="M15.5 15.5l13 13M28.5 15.5l-13 13" fill="none" stroke="#fff" stroke-width="2.8" stroke-linecap="round"/></svg>';
            case "paused":
                return '<svg viewBox="0 0 44 44" class="msla-svg-icon"><circle cx="22" cy="22" r="18" fill="#8a8886"/><rect x="16" y="14" width="4" height="16" rx="1.5" fill="#fff"/><rect x="24" y="14" width="4" height="16" rx="1.5" fill="#fff"/></svg>';
            case "cancelled":
                return '<svg viewBox="0 0 44 44" class="msla-svg-icon"><circle cx="22" cy="22" r="18" fill="#a19f9d"/><rect x="13" y="20" width="18" height="4" rx="2" fill="#fff"/></svg>';
            default: {
                const progress = this._calcProgress(kpi);
                const offOuter = CIRC_OUTER * (1 - progress);
                const offInner = CIRC_INNER * (1 - progress);
                const isN = sk === "nearing";
                const cOuter = isN ? "#F0C05A" : "#7CE5D3";
                const cInner = isN ? "#D4A030" : "#46C9B8";
                return (
                    '<svg viewBox="0 0 48 48" class="msla-svg-donut">' +
                    `<circle cx="24" cy="24" r="${R_OUTER}" class="msla-donut-track-outer"/>` +
                    `<circle cx="24" cy="24" r="${R_INNER}" class="msla-donut-track-inner"/>` +
                    `<circle cx="24" cy="24" r="${R_OUTER}" class="msla-donut-fill msla-donut-outer" data-idx="${idx}" stroke="${cOuter}" stroke-dasharray="${CIRC_OUTER}" stroke-dashoffset="${offOuter}"/>` +
                    `<circle cx="24" cy="24" r="${R_INNER}" class="msla-donut-fill msla-donut-inner" data-idx="${idx}" stroke="${cInner}" stroke-dasharray="${CIRC_INNER}" stroke-dashoffset="${offInner}"/>` +
                    "</svg>"
                );
            }
        }
    }

    // ── Timer tick ───────────────────────────────────────────

    private _tick(): void {
        if (!this._kpis.length) return;

        let needsRerender = false;
        const now = Date.now();

        for (const kpi of this._kpis) {
            if (kpi.status === SLA_STATUS.IN_PROGRESS && kpi.warningtime && now >= kpi.warningtime.getTime()) {
                kpi.status = SLA_STATUS.NEARING;
                needsRerender = true;
            }
            if (
                (kpi.status === SLA_STATUS.IN_PROGRESS || kpi.status === SLA_STATUS.NEARING) &&
                kpi.failuretime &&
                now >= kpi.failuretime.getTime()
            ) {
                kpi.status = SLA_STATUS.NONCOMPLIANT;
                needsRerender = true;
            }
        }

        if (needsRerender) {
            this._renderAll();
            return;
        }

        // Update countdowns in-place
        this._container.querySelectorAll(".msla-countdown").forEach((el) => {
            const kpi = this._kpis[parseInt(el.getAttribute("data-idx") ?? "-1", 10)];
            if (kpi) el.textContent = this._calcCountdown(kpi);
        });

        // Update donut progress in-place
        this._container.querySelectorAll(".msla-donut-outer").forEach((el) => {
            const kpi = this._kpis[parseInt(el.getAttribute("data-idx") ?? "-1", 10)];
            if (kpi) el.setAttribute("stroke-dashoffset", String(CIRC_OUTER * (1 - this._calcProgress(kpi))));
        });
        this._container.querySelectorAll(".msla-donut-inner").forEach((el) => {
            const kpi = this._kpis[parseInt(el.getAttribute("data-idx") ?? "-1", 10)];
            if (kpi) el.setAttribute("stroke-dashoffset", String(CIRC_INNER * (1 - this._calcProgress(kpi))));
        });
    }

    // ── Calculations ─────────────────────────────────────────

    private _calcProgress(kpi: KpiRecord): number {
        if (!kpi.applicablefrom || !kpi.failuretime) return 0;
        const total = kpi.failuretime.getTime() - kpi.applicablefrom.getTime();
        if (total <= 0) return 1;
        const elapsed = Date.now() - kpi.applicablefrom.getTime();
        const raw = Math.min(1, Math.max(0, elapsed / total));
        return raw > 0 ? Math.max(0.08, raw) : 0;
    }

    private _calcCountdown(kpi: KpiRecord): string {
        if (kpi.status === SLA_STATUS.SUCCEEDED) return "Completed";
        if (kpi.status === SLA_STATUS.CANCELLED) return "Cancelled";
        if (kpi.status === SLA_STATUS.PAUSED) return "Paused";

        const deadline = kpi.failuretime;
        if (!deadline) return "--:--:--";

        const diffMs = deadline.getTime() - Date.now();
        const isNeg = diffMs < 0;
        if (isNeg && !this._enableNegativeTimer && kpi.status !== SLA_STATUS.NONCOMPLIANT) return "Breached";

        const abs = Math.abs(diffMs);
        const s = Math.floor(abs / 1000);
        const d = Math.floor(s / 86400);
        const h = Math.floor((s % 86400) / 3600);
        const m = Math.floor((s % 3600) / 60);
        const sec = s % 60;
        const sign = isNeg ? "- " : "";

        if (d > 0) return `${sign}${d}d ${this._p(h)}h ${this._p(m)}m`;
        return `${sign}${this._p(h)}:${this._p(m)}:${this._p(sec)}`;
    }

    private _p(n: number): string {
        return n.toString().padStart(2, "0");
    }

    private _statusKey(kpi: KpiRecord): string {
        switch (kpi.status) {
            case SLA_STATUS.IN_PROGRESS:
                return "in-progress";
            case SLA_STATUS.NEARING:
                return "nearing";
            case SLA_STATUS.NONCOMPLIANT:
                return "noncompliant";
            case SLA_STATUS.SUCCEEDED:
                return "succeeded";
            case SLA_STATUS.PAUSED:
                return "paused";
            case SLA_STATUS.CANCELLED:
                return "cancelled";
            default:
                return "in-progress";
        }
    }

    private _badgeText(status: string): string {
        const map: Record<string, string> = {
            "in-progress": "In Progress",
            nearing: "Nearing Breach",
            noncompliant: "Breached",
            succeeded: "Succeeded",
            paused: "Paused",
            cancelled: "Cancelled",
        };
        return map[status];
    }

    private _escapeHtml(s: string): string {
        const div = document.createElement("div");
        div.textContent = s;
        return div.innerHTML;
    }
}

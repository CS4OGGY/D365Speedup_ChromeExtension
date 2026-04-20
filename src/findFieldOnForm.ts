interface RunValues {
    logicalName?: string;
    fieldLogicalName?: string;
    formName?: string;
}

interface GridOptions {
    enableSearch?: boolean; enableFilters?: boolean; enableSorting?: boolean;
    enableResizing?: boolean; showRenderTime?: boolean; showFooter?: boolean; allowHtml?: boolean;
    minSearchChars?: number; collapsed?: boolean; columnOrder?: string[] | null;
}

const KV_GRID: GridOptions = {
    allowHtml: false,
    showFooter: false,
    showRenderTime: false,
    enableSearch: false,
    enableFilters: false,
    enableSorting: false,
    enableResizing: false,
    collapsed: false,
    columnOrder: ["Key", "Value"],
};

export async function run({ logicalName, fieldLogicalName }: RunValues = {}): Promise<any> {
    const XrmCtx = (window as any).Xrm || (window as any).parent?.Xrm || (window as any).top?.Xrm;
    if (!XrmCtx) throw new Error("Xrm context not found.");

    const formContext = XrmCtx.Page;
    if (!formContext?.ui) throw new Error("This tool requires an open record form.");

    const formType: number = formContext.ui.getFormType?.() ?? 0;
    if (formType !== 1 && formType !== 2) {
        throw new Error("Please open a record (Create or Edit form) to use this tool.");
    }

    if (!logicalName?.trim()) throw new Error("Table logical name is required.");
    if (!fieldLogicalName?.trim()) throw new Error("Field logical name is required.");

    const tableName = logicalName.trim().toLowerCase();
    const fieldName = fieldLogicalName.trim().toLowerCase();

    const runtimeTable = (formContext.data.entity.getEntityName() || "").toLowerCase();
    if (runtimeTable !== tableName) {
        throw new Error(
            `Active form table is "${runtimeTable}", but input says "${tableName}". Open a "${tableName}" record first.`
        );
    }

    const currentFormItem = formContext.ui.formSelector?.getCurrentItem?.();
    const currentFormName: string = currentFormItem?.getLabel?.() || "(Current Form)";

    const t0 = performance.now();
    const occurrences: Record<string, string>[][] = [];

    (formContext.ui.tabs.get() as any[]).forEach((tab: any, tabIndex: number) => {
        const tabName: string = tab.getLabel?.() ?? tab.getName?.() ?? "";
        const tabVisible: boolean | null = tab.getVisible?.() ?? null;

        (tab.sections.get() as any[]).forEach((section: any, sectionIndex: number) => {
            const sectionName: string = section.getLabel?.() ?? section.getName?.() ?? "";
            const sectionVisible: boolean | null = section.getVisible?.() ?? null;

            (section.controls.get() as any[]).forEach((control: any, controlIndex: number) => {
                let attrName = "";
                try { attrName = (control.getAttribute?.()?.getName?.() || "").toLowerCase(); } catch (_) {}
                if (attrName !== fieldName) return;

                occurrences.push([
                    { Key: "Tab",             Value: tabName },
                    { Key: "Tab Visible",     Value: tabVisible === true ? "Yes" : "No" },
                    { Key: "Section",         Value: sectionName },
                    { Key: "Section Visible", Value: sectionVisible === true ? "Yes" : "No" },
                    { Key: "Control Name",    Value: control.getName?.() ?? "" },
                    { Key: "Control Label",   Value: control.getLabel?.() ?? "" },
                    { Key: "Field Visible",   Value: control.getVisible?.() === true ? "Yes" : "No" },
                    { Key: "Tab Position",     Value: String(tabIndex + 1) },
                    { Key: "Section Position", Value: String(sectionIndex + 1) },
                    { Key: "Control Position", Value: String(controlIndex + 1) },
                ]);
            });
        });
    });

    const fetchMs = performance.now() - t0;

    if (occurrences.length === 0) {
        return {
            __type: "interactiveTables",
            meta: { retrievedMs: Math.round(fetchMs) },
            tables: [{
                datasetName: `🔍 "${fieldName}" — not found on "${currentFormName}"`,
                gridOptions: KV_GRID,
                rows: [],
                note: `Field "<strong>${fieldName}</strong>" has no controls on the current form. It may exist on the table but is not placed on this form.`,
            }],
        };
    }

    const tables = occurrences.map((rows, i) => ({
        datasetName: `Occurrence ${i + 1} of ${occurrences.length}  —  Tab: "${rows[0].Value}"  /  Section: "${rows[2].Value}"`,
        gridOptions: KV_GRID,
        rows,
    }));

    // Prepend a summary header table
    tables.unshift({
        datasetName: `🔍 "${fieldName}" on form "${currentFormName}" — ${occurrences.length} occurrence(s)`,
        gridOptions: { ...KV_GRID, showRenderTime: true },
        rows: [
            { Key: "Table",     Value: runtimeTable },
            { Key: "Form",      Value: currentFormName },
            { Key: "Field",     Value: fieldName },
            { Key: "Found",     Value: `${occurrences.length} occurrence(s)` },
        ],
    });

    return {
        __type: "interactiveTables",
        meta: { retrievedMs: Math.round(fetchMs) },
        tables,
    };
}

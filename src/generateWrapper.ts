

interface RunValues {
    logicalName: string;
    displayName?: string;
    baseClassMode?: string;
    baseClassName?: string;
    outputComment?: string;
}

interface LocalizedLabel {
    Label?: string;
}

interface DisplayName {
    UserLocalizedLabel?: LocalizedLabel; 
}

interface AttributeMeta {
    LogicalName: string;
    AttributeType: string;
    SchemaName: string;
    DisplayName?: DisplayName;
    IsPrimaryId?: boolean;
    IsPrimaryName?: boolean;
    IsValidForRead?: boolean;
    IsValidForAdvancedFind?: boolean;
}

interface EntityMeta {
    Attributes?: AttributeMeta[];
}

interface OptionSetOption {
    Label?: {
        UserLocalizedLabel?: LocalizedLabel;
    };
    Value: number;
}

interface OptionSetResult {
    OptionSet?: {
        Options?: OptionSetOption[];
    };
}

interface GroupedAttribute {
    displayLabel: string;
    logical: string;
    type: string;
}

interface Grouped {
    primary: GroupedAttribute[];
    primaryName: GroupedAttribute[];
    strings: GroupedAttribute[];
    lookups: GroupedAttribute[];
    dates: GroupedAttribute[];
    numbers: GroupedAttribute[];
    bools: GroupedAttribute[];
    options: GroupedAttribute[];
    multiOptions: GroupedAttribute[];
    others: GroupedAttribute[];
}

export async function run(values: RunValues): Promise<string> {
    const logicalName = values.logicalName;
    const displayName = (values.displayName || logicalName).replace(/\s+/g, "") || logicalName;
    const className = displayName;
    const useDefaultBase = !values.baseClassMode || values.baseClassMode === "default";
    const baseClass = useDefaultBase ? "CRMBase" : (values.baseClassName?.trim() || "CRMBase");

    const win = window as any;
    const XrmContext = win.Xrm || win.parent?.Xrm || win.top?.Xrm;
    if (!XrmContext) {
        throw new Error("Xrm not available. Make sure this runs in a Dynamics 365 page.");
    }

    const baseUrl: string = XrmContext.Utility.getGlobalContext().getClientUrl() + "/api/data/v9.2";

    async function fetchJson(url: string): Promise<any> {
        const r = await fetch(url, { headers: { Accept: "application/json" } });
        const j = await r.json();
        if (!r.ok) throw new Error(j.error?.message || r.statusText);
        return j;
    }

    // =========================
    // 1) ENTITY METADATA
    // =========================
    const entityUrl =
        `${baseUrl}/EntityDefinitions(LogicalName='${logicalName}')` +
        `?$expand=Attributes(` +
        `$select=LogicalName,AttributeType,SchemaName,DisplayName,IsPrimaryId,IsPrimaryName,IsValidForRead,IsValidForAdvancedFind)`;

    const entityMeta: EntityMeta = await fetchJson(entityUrl);
    let attributes: AttributeMeta[] = entityMeta.Attributes || [];

    // =========================
    // 2) HELPERS
    // =========================
    const cleanName = (label: string): string =>
        (label || "")
            .replace(/\s+/g, "")
            .replace(/[^\w]/g, "")
            .replace(/^(\d)/, "_$1") || "Unnamed";

    function makeUniqueLabel(label: string, used: Set<string>): string {
        let candidate = label;
        let n = 2;
        while (used.has(candidate)) {
            candidate = `${label}_${n++}`;
        }
        used.add(candidate);
        return candidate;
    }

    // =========================
    // 3) FILTER
    // =========================
    attributes = attributes.filter((a: AttributeMeta) => {
        const ln = (a.LogicalName || "").toLowerCase();

        if (!a.IsValidForRead) return false;
        if (!a.DisplayName?.UserLocalizedLabel) return false;

        if (ln.startsWith("_")) return false;
        if (ln.endsWith("_base")) return false;

        if (a.IsPrimaryName === true) return true;
        if (ln.endsWith("yominame")) return false;

        return true;
    });

    // =========================
    // 4) OPTIONSET METADATA
    // =========================
    async function getOptionSet(logical: string, type: string): Promise<OptionSetOption[]> {
        let segment = "PicklistAttributeMetadata";
        if (type === "State") segment = "StateAttributeMetadata";
        if (type === "Status") segment = "StatusAttributeMetadata";
        if (type === "MultiSelectPicklist") segment = "MultiSelectPicklistAttributeMetadata";

        const url =
            `${baseUrl}/EntityDefinitions(LogicalName='${logicalName}')` +
            `/Attributes(LogicalName='${logical}')/Microsoft.Dynamics.CRM.${segment}?$expand=OptionSet`;

        try {
            const j: OptionSetResult = await fetchJson(url);
            return j.OptionSet?.Options || [];
        } catch {
            return [];
        }
    }

    // =========================
    // 5) GROUP ATTRIBUTES
    // =========================
    const grouped: Grouped = {
        primary: [],
        primaryName: [],
        strings: [],
        lookups: [],
        dates: [],
        numbers: [],
        bools: [],
        options: [],
        multiOptions: [],
        others: [],
    };

    const usedLabels = new Set<string>();

    for (const a of attributes) {
        const logical = a.LogicalName || "";
        const logicalLower = logical.toLowerCase();
        const labelRaw = a.DisplayName?.UserLocalizedLabel?.Label || logical;
        const labelClean = makeUniqueLabel(cleanName(labelRaw), usedLabels);
        const type = a.AttributeType;

        if (/^address\d*_addressid$/i.test(logical)) continue;
        if (
            logicalLower === "address1_addressid" ||
            logicalLower === "address2_addressid" ||
            logicalLower === "address3_addressid"
        ) {
            continue;
        }

        const isRealPrimary = a.IsPrimaryId === true || logicalLower === `${logicalName.toLowerCase()}id`;
        if (isRealPrimary && grouped.primary.length === 0) {
            grouped.primary.push({ displayLabel: `${displayName}Id`, logical, type: "Uniqueidentifier" });
            continue;
        }

        if (a.IsPrimaryName === true && grouped.primaryName.length === 0) {
            grouped.primaryName.push({ displayLabel: "Name", logical, type: "String" });
            continue;
        }

        switch (type) {
            case "String":
            case "Memo":
                grouped.strings.push({ displayLabel: labelClean, logical, type: "String" });
                break;

            case "Lookup":
            case "Customer":
            case "Owner":
                grouped.lookups.push({ displayLabel: labelClean, logical, type });
                break;

            case "DateTime":
                grouped.dates.push({ displayLabel: labelClean, logical, type });
                break;

            case "Money":
            case "Decimal":
            case "Double":
            case "Integer":
                grouped.numbers.push({ displayLabel: labelClean, logical, type });
                break;

            case "Boolean":
                grouped.bools.push({ displayLabel: labelClean, logical, type });
                break;

            case "Picklist":
            case "State":
            case "Status":
                grouped.options.push({ displayLabel: labelClean, logical, type });
                break;

            case "MultiSelectPicklist":
                grouped.multiOptions.push({ displayLabel: labelClean, logical, type });
                break;

            default:
                grouped.others.push({ displayLabel: labelClean, logical, type });
        }
    }

    // =========================
    // 6) FIELDS SECTION BUILDER
    // =========================
    function fieldsSection(title: string, arr: GroupedAttribute[]): string {
        if (!arr.length) return "";

        const lines = arr.map((a: GroupedAttribute, i: number) => {
            const isLast = i === arr.length - 1;
            const suffix = isLast ? ";" : ",";
            return `          ${a.displayLabel} = "${a.logical}"${suffix}`;
        });

        return `
          // ${title}
          public const string
${lines.join("\n")}
`;
    }

    const fieldLines =
        fieldsSection("Primary Key", grouped.primary) +
        fieldsSection("Primary Name", grouped.primaryName) +
        fieldsSection("Strings", grouped.strings) +
        fieldsSection("Lookups", grouped.lookups) +
        fieldsSection("Dates", grouped.dates) +
        fieldsSection("Currency / Numbers", grouped.numbers) +
        fieldsSection("Booleans", grouped.bools) +
        fieldsSection("Option Sets", grouped.options) +
        fieldsSection("Multi Select Option Sets", grouped.multiOptions) +
        fieldsSection("Others", grouped.others);

    // =========================
    // 7) ENUMS
    // =========================
    const enumSections: string[] = [];
    const allOptionAttrs = [...grouped.options, ...grouped.multiOptions];

    for (const a of allOptionAttrs) {
        const opts = await getOptionSet(a.logical, a.type);
        if (!opts.length) continue;

        const enumName = `${a.displayLabel}Enum`;

        const lines = opts
            .map((o: OptionSetOption) => {
                const raw = o.Label?.UserLocalizedLabel?.Label || "Unknown";
                const safe = cleanName(raw) || "Unknown";
                return `          ${safe} = ${o.Value},`;
            })
            .join("\n");

        enumSections.push(`      public enum ${enumName}\n      {\n${lines}\n      }`);
    }

    // =========================
    // 8) PROPERTY TYPES
    // =========================
    function prop(type: string, label: string): string {
        let cType = "string";

        switch (type) {
            case "Uniqueidentifier":
                cType = "Guid?";
                break;

            case "Lookup":
            case "Customer":
            case "Owner":
                cType = "EntityReference";
                break;

            case "Boolean":
                cType = "bool?";
                break;

            case "Money":
            case "Decimal":
            case "Double":
            case "Integer":
                cType = "decimal?";
                break;

            case "DateTime":
                cType = "DateTime?";
                break;

            case "Picklist":
            case "State":
            case "Status":
                cType = "OptionSetValue";
                break;

            case "MultiSelectPicklist":
                cType = "OptionSetValueCollection";
                break;
        }

        return `      public ${cType} ${label}\n      {\n          get => GetValue<${cType}>(Fields.${label});\n          set => SetValue(Fields.${label}, value);\n      }`;
    }

    function propSection(title: string, arr: GroupedAttribute[]): string {
        if (!arr.length) return "";
        return `\n      // ${title}\n${arr.map((a: GroupedAttribute) => prop(a.type, a.displayLabel)).join("\n")}`;
    }

    const propDefs =
        propSection("Primary Key", grouped.primary) +
        propSection("Primary Name", grouped.primaryName) +
        propSection("Strings", grouped.strings) +
        propSection("Lookups", grouped.lookups) +
        propSection("Dates", grouped.dates) +
        propSection("Currency / Numbers", grouped.numbers) +
        propSection("Booleans", grouped.bools) +
        propSection("Option Sets", grouped.options) +
        propSection("Multi Select Option Sets", grouped.multiOptions) +
        propSection("Others", grouped.others);

    // =========================
    // 9) BUILD OUTPUT
    // =========================
    const usings = `using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Query;`;

    const crmBase = `
  public abstract class CRMBase
  {
      protected Entity _entity;
      protected IOrganizationService _service;

      protected CRMBase(string logicalName)
      {
          _entity = new Entity(logicalName);
      }

      protected CRMBase(string logicalName, IOrganizationService service)
      {
          _entity = new Entity(logicalName);
          _service = service;
      }

      protected CRMBase(string logicalName, Guid id, IOrganizationService service)
      {
          _entity = new Entity(logicalName) { Id = id };
          _service = service;
      }

      protected CRMBase(Entity record, IOrganizationService service)
      {
          _entity = record;
          _service = service;
      }

      public object this[string fieldName]
      {
          get { return _entity.Contains(fieldName) ? _entity[fieldName] : null; }
          set { _entity[fieldName] = value; }
      }

      public T GetValue<T>(string fieldName)
      {
          return _entity.Contains(fieldName) ? (T)_entity[fieldName] : default(T);
      }

      public void SetValue(string fieldName, object value)
      {
          _entity[fieldName] = value;
      }

      public Entity ToEntity() => _entity;
      public Guid Id => _entity.Id;
      public string EntityLogicalName => _entity.LogicalName;
  }`;

    const classDef = `
  public class ${className} : ${baseClass}
  {
      #region Constructors
      protected ${className}() : base(LogicalName) { }
      protected ${className}(IOrganizationService service) : base(LogicalName, service) { }
      protected ${className}(Guid id, ColumnSet columns, IOrganizationService service)
          : base(service.Retrieve(LogicalName, id, columns), service) { }
      protected ${className}(Guid id, IOrganizationService service)
          : base(LogicalName, id, service) { }
      protected ${className}(Entity record, IOrganizationService service)
          : base(record, service) { }
      #endregion

      public static readonly string LogicalName = "${logicalName}";

      #region Constants
      public static class Fields
      {
${fieldLines}
      }
      #endregion

      #region OptionSetEnums
${enumSections.join("\n\n")}
      #endregion

      #region Fields
${propDefs}
      #endregion
  }`;

    return usings + "\n" + classDef + (useDefaultBase ? "\n" + crmBase : "");
}

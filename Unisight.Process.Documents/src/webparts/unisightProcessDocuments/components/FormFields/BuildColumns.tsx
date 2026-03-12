import * as React from "react";
import type { IColumn, IRenderFunction } from "@fluentui/react";
import type { IViewColumn, IFieldInfo } from "../../models";
import { getTaxonomyText } from "../../helpers/taxonomyDisplay";
import { Icon } from "@fluentui/react";
import { initializeFileTypeIcons, getFileTypeIconProps } from "@fluentui/react-file-type-icons";
import styles from '../UnisightProcessDocuments/UnisightProcessDocuments.module.scss';

const isDigitsOnly = (s?: string) => !!s && /^\d+$/.test(s.trim());
initializeFileTypeIcons();

export interface IBuildColumnsOptions {
  isLinksTab?: boolean;
  onOpenMenu?: (ev: React.MouseEvent<HTMLElement>, column: IColumn, fieldInfo?: IFieldInfo) => void;
  activeFilters?: string[];
}

export const buildRenderableColumns = (
  viewColumns: IViewColumn[],
  fieldInfos: Map<string, IFieldInfo>,
  options?: IBuildColumnsOptions
): IColumn[] => {
  const { 
    isLinksTab = false, 
    onOpenMenu, 
    // activeFilters 
  } = options ?? {};

  return viewColumns.map((c): IColumn => {
    const info = fieldInfos.get(c.fieldName);
    // const filtered = !!activeFilters?.[c.fieldName];

    const base: Partial<IColumn> = {
      key: c.key || c.fieldName,
      fieldName: c.fieldName,
      name: c.name || c.fieldName,
      onColumnClick: (ev, col) => {
        if (ev && onOpenMenu) {
          ev.preventDefault();
            ev.stopPropagation();
          onOpenMenu(ev as React.MouseEvent<HTMLElement>, col, info);
        }
      },
      onRenderHeader: (props, defaultRender?: IRenderFunction<any>) => {
        if (!props) return null;
        return (
          <div style={{ display: 'flex', alignItems: 'center', gap: 4 }}>
            {defaultRender?.(props)}
            <Icon iconName="Star"/>
          </div>
        );
      }
    };

    if (isLinksTab && c.fieldName === "Title") {
      return {
        ...c,
        ...base,
        onRender: (item: any) => {
          const text = item.FieldValuesAsText?.Title ?? item.Title ?? "";
          const hover =
            item.URL?.Url ??
            item.Url ??
            item.URL ??
            item.EncodedAbsUrl ??
            item.FileRef ??
            text;
          return <span title={String(hover)}>{text}</span>;
        },
      } as IColumn;
    }

    if (c.fieldName === "DocIcon") {
      return {
        ...c,
        ...base,
        minWidth: 20,
        maxWidth: 20,
        isIconOnly: true,
        onRender: (item: any) => {
          const isFolder =
            item.FSObjType === 1 ||
            item.FileSystemObjectType === 1 ||
            String(item.ContentTypeId || "").startsWith("0x0120");
          if (isFolder) return <Icon iconName="FabricFolder" />;
          const leaf =
            item.FieldValuesAsText?.FileLeafRef ??
            item.__LinkFilename ??
            "";
          const ext =
            (leaf.match(/\.([^.]+)$/)?.[1] ||
              (item.__DocIcon || "")).toString().replace(/^\./, "").toLowerCase();
          return <Icon {...getFileTypeIconProps({ extension: ext || "txt", size: 16, imageFileType: "svg" })} />;
        },
      } as IColumn;
    }

    if (["LinkFilename", "LinkFilenameNoMenu", "FileLeafRef"].includes(c.fieldName)) {
      return {
        ...c,
        ...base,
        minWidth: 100,
        maxWidth: 200,
        onRender: (item: any) => {
          const leaf = item.FieldValuesAsText?.FileLeafRef ?? item.__LinkFilename ?? "";
          const webUrl = String(item.SPWebUrl || "").replace(/\/$/, "");
          const normId = item.NormUniqueID || item.UniqueId || item.UniqueID || item.GUID || "";
          const viewerUrl = webUrl && normId
            ? `${webUrl}/_layouts/15/viewer.aspx?sourcedoc=${encodeURIComponent(normId)}`
            : (item.EncodedAbsUrl || item.FileRef || "");
          return (
            <a className={`${styles.docLinks} ${styles.links}`} href={viewerUrl} target="_blank" data-interception="off">
              {leaf}
            </a>
          );
        },
      } as IColumn;
    }

    if (!info) {
      return { ...c, ...base } as IColumn;
    }

    const t = info.TypeAsString;

    if (t === "TaxonomyFieldType" || t === "TaxonomyFieldTypeMulti") {
      return {
        ...c,
        ...base,
        minWidth: 100,
        maxWidth: 150,
        onRender: (item: any) => {
          const txt = getTaxonomyText(item, c.fieldName);
          const fallback = item.__TaxoResolved?.[c.fieldName];
          return <span>{txt || fallback || ""}</span>;
        },
      } as IColumn;
    }

    if (t === "User" || t === "UserMulti") {
      return {
        ...c,
        ...base,
        minWidth: 100,
        maxWidth: 150,
        onRender: (item: any) => {
          const fvat = item.FieldValuesAsText?.[c.fieldName];
          if (typeof fvat === "string" && !isDigitsOnly(fvat)) return <span>{fvat}</span>;
          const val = item[c.fieldName];
          if (!val) return <span />;
          if (Array.isArray(val)) {
            const labels = val.map(v => v?.Title ?? v?.title ?? "").filter(Boolean);
            return <span>{labels.join("; ")}</span>;
          }
          return <span>{val?.Title ?? val?.title ?? ""}</span>;
        },
      } as IColumn;
    }

    if (t === "Lookup" || t === "LookupMulti") {
      const display = info.LookupField || "Title";
      return {
        ...c,
        ...base,
        onRender: (item: any) => {
          const fvat = item.FieldValuesAsText?.[c.fieldName];
          if (typeof fvat === "string" && fvat) return <span>{fvat}</span>;
          const val = item[c.fieldName];
          if (!val) return <span />;
          if (Array.isArray(val)) {
            return <span>{val.map(v => v?.[display] ?? "").filter(Boolean).join("; ")}</span>;
          }
          return <span>{val?.[display] ?? ""}</span>;
        },
      } as IColumn;
    }

    if (t === "DateTime") {
      return {
        ...c,
        ...base,
        minWidth: 40,
        maxWidth: 100,
        onRender: (item: any) => {
          const raw = item.FieldValuesAsText?.[c.fieldName] ?? item[c.fieldName];
          if (!raw) return <span />;
          const d = new Date(raw);
          // Format as yyyy-MM-DDThh:mm:ssZ
          const iso = isNaN(d.getTime()) ? "" : d.toISOString().replace(/\.\d{3}Z$/, "Z");
          return <span alt-text={iso}>{isNaN(d.getTime()) ? String(raw) : d.toLocaleDateString()}</span>;
        },
      } as IColumn;
    }

    return {
      ...c,
      ...base,
      minWidth: 100,
      maxWidth: 150,
      onRender: (item: any) =>
        <span>{item.FieldValuesAsText?.[c.fieldName] ?? item[c.fieldName] ?? ""}</span>,
    } as IColumn;
  });
};
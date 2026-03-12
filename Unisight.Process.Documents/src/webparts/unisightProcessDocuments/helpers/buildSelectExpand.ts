import type { IColumn } from '@fluentui/react';
import type { IFieldInfo } from '../models';
import { 
  toODataName, 
  // isUnderscoreField 
} from './odataNames';

export const buildSelectExpand = (
  viewColumns: IColumn[],
  fieldInfos: Map<string, IFieldInfo>
): { select: string[]; expand: string[] } => {
  const select = new Set<string>([
    "Id",
    "FileRef",
    "FileLeafRef",
    "FieldValuesAsText",
  ]);
  const expand = new Set<string>([
    "FieldValuesAsText",
  ]);

  for (const col of viewColumns) {
    const internal = (col as any).fieldName as string | undefined;
    if (!internal) continue;

    const odata = toODataName(internal);
    // Select both the raw value and its text equivalent
    select.add(odata);
    select.add(`FieldValuesAsText/${odata}`);

    // If you have special handling for types, you can use fieldInfos.get(internal) here
    // For most underscore pseudo-fields (like _UIVersionString), treating as text is fine.
  }

  return { select: Array.from(select), expand: Array.from(expand) };
};
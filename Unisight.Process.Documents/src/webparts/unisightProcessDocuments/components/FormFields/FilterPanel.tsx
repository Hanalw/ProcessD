import * as React from 'react';
import { useEffect, useMemo, useState, useCallback, FC } from 'react';
import { Panel } from '@fluentui/react/lib/Panel';
import { Checkbox, PrimaryButton } from '@fluentui/react';
import { IFieldInfo } from '../../models';
import * as strings from 'UnisightProcessDocumentsWebPartStrings';
import { IFilterItem } from '../UnisightProcessDocuments';

import styles from '../UnisightProcessDocuments/UnisightProcessDocuments.module.scss';

// Type for values (taxonomy and non-taxonomy)
type TaxoOption = { label: string; value: string };
type FilterPanelValue = string | TaxoOption;

interface FilterPanelProps {
  isOpen: boolean;
  title?: string;
  onDismiss?: () => void;
  fieldName: string | undefined
  fieldInfo: IFieldInfo | null;
  values: FilterPanelValue[]; // now supports both string and {label, value}
  selectedValues?: IFilterItem[];              // new (optional controlled values)
  onSelectionChange?: (values: IFilterItem[]) => void; // new callback
  onSelectedValuesChange?: (values: string[]) => void;
  searchDocs: any[] | undefined;
  onApply?: () => void;
  onClear?: (fieldName?: string) => void;
  useFilters: any[];
}

const FilterPanel: FC<FilterPanelProps> = (props) => {
  const {
    isOpen,
    fieldName,
    title = `${strings.filter}: ${fieldName}`,
    onDismiss,
    values,
    // fieldInfo,
    selectedValues,
    onSelectionChange,
    onClear,
    onApply,
    useFilters
  } = props;

  // local selection state now uses IFilterItem[]
  const [, setSelected] = useState<IFilterItem[]>(selectedValues ?? []);
  

  const formattedEntries = useMemo<{ label: string; value: string }[]>(() => {
    return values.map((v) =>
      typeof v === "string"
        ? { label: v, value: v }
        : { label: v.label, value: v.value }
    );
  }, [values]);

  // keep local state in sync if parent controls it
  useEffect(() => {
    console.log(values)
    if (selectedValues) setSelected(selectedValues);
  }, [selectedValues]);

  // NEW: clear all checkboxes if useFilters is empty
  useEffect(() => {
    if (!useFilters || useFilters.length === 0) {
      setSelected([]);
    }
  }, [useFilters]);

  // Notify parent (for legacy support)
  const notifyParent = useCallback((items: IFilterItem[]) => {
    if (props.onSelectedValuesChange) {
      const flat = items.reduce<string[]>((acc, i) => {
        if (i.value) acc.push(...i.value);
        return acc;
      }, []);
      props.onSelectedValuesChange(flat);
    }
  }, [props.onSelectedValuesChange]);

  // Toggle selection of a value (TermGuid or plain value)
  const toggle = (value: string) => {
    const column = fieldName ?? '';
    setSelected(prev => {
      let cleared = false;
      const next = [...prev];
      const idx = next.findIndex(it => it.columnName === column);

      if (idx === -1) {
        // first value for this column
        next.push({ columnName: column, value: [value] });
      } else {
        const currentValues = next[idx].value;
        if (currentValues.includes(value)) {
          // remove this value
          const filtered = currentValues.filter(v => v !== value);
          if (filtered.length === 0) {
            next.splice(idx, 1);
            cleared = true;
          } else {
            next[idx] = { ...next[idx], value: filtered };
          }
        } else {
          // add value
          next[idx] = { ...next[idx], value: [...currentValues, value] };
        }
      }

      onSelectionChange?.(next);
      notifyParent(next);

      // invoke onClear when last value removed
      if (cleared) {
        onClear?.();
      }
      return next;
    });
  };

  // Clear handler
  const handleClear = () => {
    setSelected([]);
    onSelectionChange?.([]);
    notifyParent([]);
    onClear?.(fieldName); // Pass which filter to clear
  };

  // checks if checkbox is included in useFilters, if yes keep that checkbox checked on rerender of panel
  const isChecked = (value: string): boolean => {
    const filter = useFilters.find(
      (item: any) => item.fieldName === fieldName
    );
    return Array.isArray(filter?.values) && filter.values.includes(value);
  };

  // Determine if any checkboxes are checked for this field when panel is open
  const anyChecked = useMemo(() => {
    const filter = useFilters.find(
      (item: any) => item.fieldName === fieldName
    );
    return Array.isArray(filter?.values) && filter.values.length > 0;
  }, [useFilters, fieldName]);

  const onRenderFooterContent = () => (
    <div>
      <PrimaryButton text={'Apply'} onClick={onApply} className={styles.buttons}/>
      <PrimaryButton text={'Clear'} onClick={handleClear} disabled={!anyChecked} />
    </div>
  );

  return (
    <Panel
      isOpen={isOpen}
      onDismiss={onDismiss}
      headerText={title}
      closeButtonAriaLabel="Close"
      onRenderFooterContent={onRenderFooterContent}
      className={styles.filterpanel}
    >
      {formattedEntries.map(entry => (
        <Checkbox
          key={entry.value}
          label={entry.label}
          checked={isChecked(entry.value)}
          onChange={() => toggle(entry.value)}
          className={styles.checkboxes}
        />
      ))}
    </Panel>
  );
};

export default FilterPanel;
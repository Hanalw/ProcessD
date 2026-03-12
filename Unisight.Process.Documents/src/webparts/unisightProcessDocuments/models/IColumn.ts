export interface IColumn {
  key: string,
  name: string | undefined,
  fieldName: string,
  minWidth: number,
  maxWidth: number,
  isResizable: boolean,
  data: string,
  isSorted: boolean,
  isSortedDescending: boolean,
  isRowHeader: boolean,
  iconName: string,
  isIconOnly: boolean,
}
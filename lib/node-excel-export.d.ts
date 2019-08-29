declare module 'node-excel-export' {

  type CellStyle = {
    fill?: {
      fgColor: {
        rgb: string;
      }
    },
    font?: {
      color?: {
        rgb: string;
      },
      sz?: number;
      bold?: boolean;
      underline?: boolean;
    },
    alignment?: {
      horizontal?: 'left' | 'center' | 'right',
      vertical?: 'top' | 'center' | 'bottom',
    },
  };

  type Heading<TRowData> = [{
    value: keyof TRowData;
    style: CellStyle
  }| keyof TRowData[]][]

  type Merges = {
    start: {
      row: number;
      column: number;
    },
    end: {
      row: number;
      column: number;
    };
  }[]

  type Specification<TRowData> = {
    [CellName in keyof TRowData]: {
      displayName: string;
      headerStyle?: ((value: TRowData[CellName], row: TRowData) => CellStyle) | CellStyle;
      cellStyle?: ((value: TRowData[CellName], row: TRowData) => CellStyle) | CellStyle;
      cellFormat?: (value: TRowData[CellName], row: TRowData) => string;
      width: string | number;
    }
  }

  function buildExport<TRowData>(sheets: {
    name: string;
    heading: Heading<TRowData> | undefined;
    merges: Merges | undefined;
    specification: Specification<TRowData>;
    data: TRowData[];
  }[]): Buffer;

}
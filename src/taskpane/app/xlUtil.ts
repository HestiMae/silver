import XLSX, { CellObject, WorkSheet } from "xlsx";

export default class xlUtil {

  static ec = (r, c) => XLSX.utils.encode_cell({ r: r, c: c });
  static er = (s, e) => XLSX.utils.encode_range(s, e);
  static erc = (sr, sc, er, ec) => XLSX.utils.encode_range({ r: sr, c: sc }, { r: er, c: ec });
  static used_range = (s) => XLSX.utils.decode_range(s["!ref"]);

  static filter_sheet(sheet: WorkSheet, header: string, filterValue: string): number {
    let range = this.used_range(sheet)
    let column = this.range_to_string_array(sheet, this.header_search(sheet, range, "station"), 1);
    let filter_rows: Array<number> = [];
    column.forEach((value, i) => {if (value != filterValue) {filter_rows.push(range.s.r + i + 1)}})
    if (filter_rows.length > 0) {
      this.delete_many_rows(sheet, filter_rows)
    }
    return filter_rows.length
  }

  static delete_many_rows(sheet: WorkSheet, row_index: Array<number>) {
    let range = this.used_range(sheet);
    let deletedRows = 0;
    for (let row = range.s.r; row <= range.e.r; row++) {
      if (row == row_index[deletedRows]) {
        deletedRows += 1
      }
      else if (deletedRows > 0) {
        // Overwrite row
        for (let col = range.s.c; col <= range.e.c; col++) {
          sheet[this.ec(row - deletedRows, col)] = sheet[this.ec(row, col)];
        }
      }
    }
    sheet["!ref"] = this.erc(range.s.r, range.s.c, range.e.r - deletedRows, range.e.c);
  }

  static get_cell_value(cell: CellObject): string {
    if (cell == undefined) {
      return "";
    }
    return cell.v.toString();
  }

  static get_cell_value_numeric(cell: CellObject): number
  {
    if (cell == undefined || cell.t != "n")
    {
      return -1;
    }
    return cell.v as number
  }

  static header_search(sheet: XLSX.WorkSheet, range: XLSX.Range, header: string): XLSX.Range {
    return this.header_search_multi(sheet, range, [header])[0]
  }

  static header_search_multi(sheet: XLSX.WorkSheet, range: XLSX.Range, header: Array<string>): Array<XLSX.Range> {
    let outArray = new Array<XLSX.Range>(header.length)
    for (let column = range.s.c; column <= range.e.c; column++) {
      let cell: XLSX.CellObject = sheet[this.ec(range.s.r, column)];
      let index = header.findIndex(value => value == this.get_cell_value(cell));
      if (index != -1) {
        outArray[index] = { s: { r: range.s.r, c: column }, e: { r: range.e.r, c: column } };
      }
    }
    return outArray
  }

  static range_to_string_array(sheet: XLSX.WorkSheet, range: XLSX.Range, offset: number): Array<string> {
    let rangeArray = [];
    for (let row = range.s.r + offset; row <= range.e.r; row++) {
      let cell: XLSX.CellObject = sheet[this.ec(row, range.s.c)];
      rangeArray.push(this.get_cell_value(cell));
    }
    return rangeArray;
  }
  static range_to_number_array(sheet: XLSX.WorkSheet, range: XLSX.Range, offset: number): Array<number> {
    let rangeArray = [];
    for (let row = range.s.r + offset; row <= range.e.r; row++) {
      let cell: XLSX.CellObject = sheet[this.ec(row, range.s.c)];
      rangeArray.push(this.get_cell_value_numeric(cell));
    }
    return rangeArray;
  }

  static common_prefix(files: FileList) {
    let prefix = null;
    Array.prototype.forEach.call(files, (file) => prefix = (prefix) ? this.shared_prefix(prefix, file.name) : file.name);
    return prefix;
  }

  static shared_prefix(str1: string, str2: string) {
    let prefix = "";
    for (let i = 0; i < Math.min(str1.length, str2.length); i++) {
      if (str1[i] == str2[i]) {
        prefix += str1[i];
      } else {
        break;
      }
    }
    return prefix;
  }
}
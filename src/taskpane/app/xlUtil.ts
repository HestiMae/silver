import XLSX, { CellObject, WorkSheet } from "xlsx";

export default class xlUtil {

  static ec = (r, c) => XLSX.utils.encode_cell({ r: r, c: c });

  static filter_sheet(sheet: WorkSheet, header: string, filterValue: string) {
    let range = XLSX.utils.decode_range(sheet["!ref"])
    let column = this.range_to_array(sheet, this.header_search(sheet, range, "station"), 1);
    let filter_rows: Array<number> = [];
    column.forEach((value, i) => {if (value != filterValue) {filter_rows.push(range.s.r + i + 1)}})
    console.log("Matching Rows: " + (column.length - filter_rows.length))
    if (filter_rows.length > 0) {
      this.delete_many_rows(sheet, filter_rows)
    }
  }

  static delete_many_rows(sheet: WorkSheet, row_index: Array<number>) {
    let range = XLSX.utils.decode_range(sheet["!ref"]);
    let deletedRows = 0;
    console.log(row_index)
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
    sheet["!ref"] = XLSX.utils.encode_range(range.s, {r:range.e.r - deletedRows - 1, c:range.e.c});
    console.log("Deleted " + deletedRows.toString() + " rows - old end " + range.e.r + " new end" + (range.e.r - deletedRows))
  }

  static get_cell_value(cell: CellObject): string {
    if (cell == undefined) {
      return "";
    }
    return cell.v.toString();
  }

  static header_search(sheet: XLSX.WorkSheet, range: XLSX.Range, header: string): XLSX.Range {
    for (let column = range.s.c; column < range.e.c; column++) {
      let cell: XLSX.CellObject = sheet[XLSX.utils.encode_cell({ r: range.s.r, c: column })];
      if (cell.v.toString() == header) {
        return { s: { r: range.s.r, c: column }, e: { r: range.e.r, c: column } };
      }
    }
  }

  static range_to_array(sheet: XLSX.WorkSheet, range: XLSX.Range, offset: number): Array<string> {
    let rangeArray = [];
    for (let row = range.s.r + offset; row < range.e.r; row++) {
      let cell: XLSX.CellObject = sheet[XLSX.utils.encode_cell({ r: row, c: range.s.c })];
      rangeArray.push(this.get_cell_value(cell));
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
import { Component } from "@angular/core";
// images references in the manifest
import "../../../assets/icon-16.png";
import "../../../assets/icon-32.png";
import "../../../assets/icon-80.png";

import fileDialog from "file-dialog";
import XLSX, { WorkBook } from "xlsx";
import { BrowserModule } from "@angular/platform-browser";
import { FormControl } from "@angular/forms";
import Worksheet = Excel.Worksheet;
/* global console, Excel, require */

interface Station {
  id: string;
  name: string;
}


@Component({
  selector: "app-home",
  templateUrl: "src/taskpane/app/app.component.html",
})
export default class AppComponent {
  stations: Station[] = [{id:"TestID", name:"testName"}]
  welcomeMessage = "Silver";
  mergeWorkbook : WorkBook;


  async select_files() {
    try {
      fileDialog({multiple:true}).then(files => {
        // Create New Workbook
        this.mergeWorkbook = XLSX.utils.book_new();
        let prefix = AppComponent.common_prefix(files)
        // Parse all CSV files
        Promise.all(Array.prototype.map.call(files, (file) => file.text())).then((file_texts) => {
          // For each file data
          file_texts.forEach(((text, i) => {
            // Read into workbook and transfer sheet
            let tempWorkbook = XLSX.read(text, { type: "binary" });
            XLSX.utils.book_append_sheet(this.mergeWorkbook, tempWorkbook.Sheets[tempWorkbook.SheetNames[0]], files[i].name.substring(prefix.length, 30+prefix.length).split('.')[0])
          }))
          // Populate
          let sheet : XLSX.WorkSheet = this.mergeWorkbook.Sheets[this.mergeWorkbook.SheetNames[0]];
          let range = XLSX.utils.decode_range(sheet["!ref"])
          this.stations = []
          for (let row = range.s.r + 1; row < range.e.r; row++) {
            let cell: XLSX.CellObject = sheet[XLSX.utils.encode_cell({c:range.s.c, r:row})]
            this.stations.push({id:cell.v.toString(), name: cell.v.toString()})
          }
        })
      });
    } catch (error) {
      console.error(error)
    }
  }

  static common_prefix(files: FileList) {
    let prefix = null;
    Array.prototype.forEach.call(files, (file) => prefix = (prefix) ? this.shared_prefix(prefix, file.name) : file.name)
    return prefix
  }

  static shared_prefix(str1: string, str2: string) {
    let prefix = "";
    for (let i = 0; i < Math.min(str1.length, str2.length); i++) {
      if (str1[i] == str2[i]) {
        prefix += str1[i]
      }
      else {
        break;
      }
    }
    return prefix
  }

  async save_workbook() {
    try {
      XLSX.writeFile(this.mergeWorkbook, "merged.xlsx");
    } catch (error) {
      console.error(error)
    }
  }

  async run() {
    try {
      await Excel.run(async (context) => {
        /**
         * Excel API here
         */
        const range = context.workbook.getSelectedRange();

        // Read the range address
        range.load("address");

        // Update the fill color
        // range.format.fill.color = "yellow";

        await context.sync();
        console.log(`The range address was ${range.address}.`);
      });
    } catch (error) {
      console.error(error);
    }
  }
}

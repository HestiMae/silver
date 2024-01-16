import { Component } from "@angular/core";
// images references in the manifest
import "../../../assets/icon-16.png";
import "../../../assets/icon-32.png";
import "../../../assets/icon-80.png";

const template = require("./app.component.html");
import fileDialog from "file-dialog";
import XLSX, { WorkBook } from "xlsx";
/* global console, Excel, require */

@Component({
  selector: "app-home",
  template
})
export default class AppComponent {
  welcomeMessage = "Welcome";

  async run() {
    try {
      /**
       * Sheetjs here
       */
      fileDialog({multiple:true}).then(files => {
        const wb = XLSX.utils.book_new();
        Promise.all(Array.prototype.map.call(files, (file) => file.text())).then((file_texts) => {
          file_texts.forEach(((text, i) => {
            let tempWorkbook = XLSX.read(text, { type: "binary" });
            XLSX.utils.book_append_sheet(wb, tempWorkbook.Sheets[tempWorkbook.SheetNames[0]], files[i].name.substring(0, 30).split('.')[0])

          }))
          XLSX.writeFile(wb, "merged.xlsx");
        })
      });


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

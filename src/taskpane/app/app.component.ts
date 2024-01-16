import { Component, OnInit, ViewChild } from "@angular/core";
// images references in the manifest
import "../../../assets/icon-16.png";
import "../../../assets/icon-32.png";
import "../../../assets/icon-80.png";

import fileDialog from "file-dialog";
import xlUtil from "./xlUtil";
import XLSX from "xlsx";
import ready from "document-ready"

/* global console, Excel, require */

interface Station {
  id: string;
  name: string;
}


@Component({
  selector: "app-home",
  templateUrl: "src/taskpane/app/app.component.html"
})
export default class AppComponent implements OnInit{
  @ViewChild('stationSelect') stationSelect: ViewChild
  stations: Station[];
  welcomeMessage = "Silver";
  mergeWorkbook: XLSX.WorkBook;
  stationSelected: string;
  disableSelect: boolean;

  ngOnInit() {
    this.stationSelected = "[Export All]";
    this.disableSelect = true;
  }

  async select_files() {
    try {
      fileDialog({ multiple: true }).then(files => {
        // Create New Workbook
        this.mergeWorkbook = XLSX.utils.book_new();
        let prefix = xlUtil.common_prefix(files);
        // Parse all CSV files
        Promise.all(Array.prototype.map.call(files, (file) => file.text())).then((file_texts) => {
          // For each file data
          file_texts.forEach(((text, i) => {
            // Read into workbook and transfer sheet
            let tempWorkbook = XLSX.read(text, { type: "binary" });
            XLSX.utils.book_append_sheet(this.mergeWorkbook, tempWorkbook.Sheets[tempWorkbook.SheetNames[0]], files[i].name.substring(prefix.length, 30 + prefix.length).split(".")[0]);
          }));
          // Populate
          let sheet: XLSX.WorkSheet = this.mergeWorkbook.Sheets["station_audit"];
          let idRange = xlUtil.header_search(sheet, XLSX.utils.decode_range(sheet["!ref"]), "station"); // Keyed by name rn
          let idArray = xlUtil.range_to_array(sheet, idRange, 1);

          let nameRange = xlUtil.header_search(sheet, XLSX.utils.decode_range(sheet["!ref"]), "station");
          let nameArray = xlUtil.range_to_array(sheet, nameRange, 1);

          this.stations = [{ id: "[Export All]", name: "[Export All]" }];
          this.stations = this.stations.concat(idArray.map((element, i) => ({
            id: element,
            name: nameArray[i]
          })).sort((station1, station2) => {
            if (station1.name > station2.name) {
              return 1;
            }
            if (station1.name < station2.name) {
              return -1;
            }
            return 0;
          }));

          this.disableSelect = false;
        });
      });
    } catch (error) {
      console.error(error);
    }
  }


  async save_workbook() {
    try {
      // Filter by Selection
      if (this.mergeWorkbook) {
        if (this.stationSelected != "[Export All]") this.mergeWorkbook.SheetNames.forEach((name) => xlUtil.filter_sheet(this.mergeWorkbook.Sheets[name], "station", this.stationSelected));
        XLSX.writeFile(this.mergeWorkbook, this.stationSelected +".xlsx");
      }
    } catch (error) {
      console.error(error);
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

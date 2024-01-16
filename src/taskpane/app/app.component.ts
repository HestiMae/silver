import { Component, OnInit, ViewChild } from "@angular/core"

import fileDialog from "file-dialog"
import XLSX from "xlsx"

import SilverField from "./SilverField"
import xlUtil from "./xlUtil"
import wordUtil from "./wordUtil"

/* global console, Excel, require */

interface Station {
  id: string;
  name: string;
}

const fields_json = require("./resources/fields.json")
@Component({
  selector: "app-home",
  templateUrl: "src/taskpane/app/app.component.html"
})
export default class AppComponent implements OnInit {
  @ViewChild("stationSelect") stationSelect: ViewChild
  stations: Station[]
  fields: Array<SilverField> = []
  mergeWorkbook: XLSX.WorkBook
  stationSelected: string
  newDocumentBase64: string
  statusBarText: string


  ngOnInit() {
    this.reset_workbook()
  }

  reset_workbook() {
    this.fields = undefined
    this.mergeWorkbook = undefined
    this.stations = undefined
    this.stationSelected = "Import Data First"
    this.set_status(`Workbook Reset`)
  }

  load_fields() {
    this.fields = SilverField.fromJSON(fields_json)
    this.set_status(`${this.fields.length} Field(s) Loaded`)
  }

  async import_csvs() {
    if (!this.mergeWorkbook) {
      try {
        fileDialog({ multiple: true }).then(files => {
          // Create New Workbook
          this.mergeWorkbook = XLSX.utils.book_new()
          let prefix = xlUtil.common_prefix(files)
          // Parse all CSV files
          Promise.all(Array.prototype.map.call(files, (file) => file.text())).then((file_texts) => {
            // For each file data
            file_texts.forEach(((text, i) => {
              // Read into workbook and transfer sheet
              let tempWorkbook = XLSX.read(text, { type: "binary" })
              XLSX.utils.book_append_sheet(this.mergeWorkbook, tempWorkbook.Sheets[tempWorkbook.SheetNames[0]], files[i].name.substring(prefix.length, 30 + prefix.length).split(".")[0])
            }))
            this.populate_stations()
          })
        })
      } catch (error) {
        console.error(error)
      }
    }
  }

  import_xlsx() {
    if (!this.mergeWorkbook) {
      fileDialog().then(files => {
        Promise.all(Array.prototype.map.call(files, (file) => file.text())).then((file_texts) => {
          this.mergeWorkbook = XLSX.read(file_texts[0], { type: "binary" })
        })
      })
    }
  }

  set_status(text) {
    this.statusBarText = text
    console.log(text)
  }

  populate_stations() {
    if (this.mergeWorkbook) {
      let sheet: XLSX.WorkSheet = this.mergeWorkbook.Sheets["station_audit"]
      let idRange = xlUtil.header_search(sheet, XLSX.utils.decode_range(sheet["!ref"]), "station") // Keyed by name rn
      let idArray = xlUtil.range_to_string_array(sheet, idRange, 1)

      let nameRange = xlUtil.header_search(sheet, XLSX.utils.decode_range(sheet["!ref"]), "station")
      let nameArray = xlUtil.range_to_string_array(sheet, nameRange, 1)

      this.stations = [{ id: "[Export All]", name: "[Export All]" }]
      this.stations = this.stations.concat(idArray.map((element, i) => ({
        id: element,
        name: nameArray[i]
      })).sort((station1, station2) => {
        if (station1.name > station2.name) {
          return 1
        }
        if (station1.name < station2.name) {
          return -1
        }
        return 0
      }))

      this.set_status(`Export Loaded with ${this.stations.length} stations and ${this.mergeWorkbook.SheetNames.length} sheets`)
    }
  }

  async filter_workbook() {
    if (this.mergeWorkbook) {
      let filter_rows = 0
      if (this.stationSelected != "[Export All]") {
        this.mergeWorkbook.SheetNames.forEach((name) => filter_rows += xlUtil.filter_sheet(this.mergeWorkbook.Sheets[name], "station", this.stationSelected))
      }
      this.set_status(`Filtered Export to '${this.stationSelected}', removing ${filter_rows} rows`)
      this.stations = undefined
    }
  }

  async save_workbook() {
    if (this.mergeWorkbook) {
      try {
        XLSX.writeFile(this.mergeWorkbook, this.stationSelected + ".xlsx")
        this.set_status(`Saved Export to ${this.stationSelected}.xlsx`)
      } catch (error) {
        console.error(error)
        this.set_status(`WRITE ERROR: Failed to save export`)
      }
    }
  }

  async extract_data() {
    if (this.mergeWorkbook && !this.stations) {
      this.load_fields()
      SilverField.extract_field_ranges(this.fields, this.mergeWorkbook)
      this.set_status(`Found Data for ${this.fields.filter((field) => field.outValue != undefined).length}/${this.fields.length} fields`)
    }
  }

  async set_binding() {
    if (this.fields) {
      Promise.all(this.fields.map((field) => field.output.bind())).then(async (bindSuccesses) => {
        this.set_status(`Successfully bound ${bindSuccesses.filter(Boolean).length}/${bindSuccesses.filter((result) => result != null).length} fields.`)
        let asyncSet = this.fields.map((field) => field.output.outputAsync(field.outValue))
        this.statusBarText += ` Set Data for ${asyncSet.filter(Boolean).length}/${asyncSet.filter((result) => result != null).length} text fields `
        wordUtil.wordResultPromise((context: Word.RequestContext) => Promise.all(this.fields.map((field) => field.output.outputWord(context, field.outValue)))).then((setSuccesses) => {
          this.statusBarText += `and ${setSuccesses.filter(Boolean).length}/${setSuccesses.filter((result) => result != null).length} other fields`
        })
      })
    }
  }

  async pick_template() {
    fileDialog().then(files => {
      this.loadfileBase64(files[0]).then(base64 => {
        this.newDocumentBase64 = base64
      })
    })
  }

  async loadFileName(document: Office.Document) {
    return new Promise((resolve) => {
      document.getFilePropertiesAsync(null, (res) => {
        resolve(res && res.value && res.value.url ? res.value.url : "")
      })
    })
  }

  async loadfileBase64(file: File): Promise<string> {
    return new Promise<string>((resolve, reject) => {
      const reader = new FileReader()
      reader.readAsDataURL(file)
      reader.onload = () => resolve(reader.result.toString().replace(/^data:(.*,)?/, "") + "=".repeat(4 - (reader.result.toString().replace(/^data:(.*,)?/, "").length % 4)))
      reader.onerror = error => reject(error)
    })
  }

  async save_document() {
    try {
      await Word.run(async (context) => {
        let newDocument: Word.DocumentCreated = context.application.createDocument(this.newDocumentBase64)
        //newDocument.properties.customProperties.getItem("AF_StationName").value = this.stationSelected
        let nameParagraph = newDocument.body.insertParagraph((this.stationSelected + "_REV0"), Word.InsertLocation.start)
        newDocument.save()
        nameParagraph.delete()
        newDocument.open()
        await context.sync().catch((e) => console.error(e))
      })
    } catch (error) {
      console.error(error)
    }
  }
}

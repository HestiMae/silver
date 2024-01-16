import { Component, OnInit, ViewChild } from "@angular/core"

import fileDialog from "file-dialog"
import XLSX from "xlsx"

import SilverField from "./SilverField"
import xlUtil from "./xlUtil"
import { TypedJSON } from "typedjson"

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
  welcomeMessage = "Silver"
  mergeWorkbook: XLSX.WorkBook
  stationSelected: string
  disableSelect: boolean
  newDocumentBase64: string
  field_stationName: string
  statusBarText: string


  ngOnInit() {
    this.stationSelected = "[Export All]"
    this.disableSelect = true
    Array.prototype.forEach.call(fields_json, (field) => this.fields.push(TypedJSON.parse(field, SilverField)))
    this.statusBarText = `${this.fields.length} Field(s) Loaded`
  }

  async select_files() {
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
          // Populate
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

          this.disableSelect = false
          this.statusBarText = `Export Loaded with ${this.stations.length} stations and ${this.mergeWorkbook.SheetNames.length} sheets`
        })
      })
    } catch (error) {
      console.error(error)
    }
  }

  async filter_workbook() {
    if (this.mergeWorkbook) {
      if (this.stationSelected != "[Export All]") {
        let filter_rows = 0
        this.mergeWorkbook.SheetNames.forEach((name) => filter_rows += xlUtil.filter_sheet(this.mergeWorkbook.Sheets[name], "station", this.stationSelected))
        this.statusBarText = `Filtered Export to '${this.stationSelected}', removing ${filter_rows} rows`
        this.disableSelect = true
      }
    }
  }

  async save_workbook() {
    try {
      // Filter by Selection
      if (this.mergeWorkbook) {
        XLSX.writeFile(this.mergeWorkbook, this.stationSelected + ".xlsx")
        this.statusBarText = `Saved Export to ${this.stationSelected}.xlsx`
      }
    } catch (error) {
      console.error(error)
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

  async initialize_binding() {
    this.statusBarText = `Successfully bound ${this.fields.filter((field) => field.output.bind()).length}/${this.fields.length} fields`
  }

  async set_binding() {
    this.fields.forEach((field) => field.output.outputAsync(field.outValue))
    try {
      await Word.run(async (context) => {
        this.fields.forEach((field) => field.output.outputWord(context, field.outValue))
      })
    } catch (error) {
      console.error(error)
    }
    this.statusBarText = `Set Bindings`
  }

  async save_document() {
    try {
      await Word.run(async (context) => {
        let newDocument: Word.DocumentCreated = context.application.createDocument(this.newDocumentBase64)
        //newDocument.properties.customProperties.getItem("AF_StationName").value = this.stationSelected
        let nameParagraph = newDocument.body.insertParagraph((this.field_stationName + "_REV0"), Word.InsertLocation.start)
        newDocument.save()
        nameParagraph.delete()
        newDocument.open()
        await context.sync().catch((e) => console.error(e))
      })
    } catch (error) {
      console.error(error)
    }
  }

  async extract_data() { //TODO: one line headers
    this.mergeWorkbook.SheetNames.forEach((sheetName) => {
      let sheet = this.mergeWorkbook.Sheets[sheetName]
      let matchedFields = this.fields.filter((field) => field.sheet_name == sheetName)
      let headers: Array<string> = []
      matchedFields.forEach((match) => headers.push(match.field_name))
      let headerSearchResult = xlUtil.header_search_multi(sheet, xlUtil.used_range(sheet), headers)
      headerSearchResult.forEach((range, i) => matchedFields[i].set_data(range, sheet))
    })
    this.statusBarText = `Found Data for ${this.fields.filter((field) => field.dataRange).length}/${this.fields.length} fields`
  }
}

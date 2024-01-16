import {Component, OnInit, ViewChild} from "@angular/core";
// images references in the manifest
import "../../../assets/icon-16.png";
import "../../../assets/icon-32.png";
import "../../../assets/icon-80.png";

import fileDialog from "file-dialog";
import xlUtil from "./xlUtil";
import XLSX from "xlsx";
import DocumentCreated = Word.DocumentCreated;
import {Base64} from "js-base64";
import InsertLocation = Word.InsertLocation;

/* global console, Excel, require */

interface Station {
    id: string;
    name: string;
}


@Component({
    selector: "app-home",
    templateUrl: "src/taskpane/app/app.component.html"
})
export default class AppComponent implements OnInit {
    @ViewChild('stationSelect') stationSelect: ViewChild
    stations: Station[];
    welcomeMessage = "Silver";
    mergeWorkbook: XLSX.WorkBook;
    stationSelected: string;
    disableSelect: boolean;
    newDocumentBase64: string;
    field_stationName: string;
    NameBinding: Office.TextBinding;
    VITBinding: Office.TableBinding;


    ngOnInit() {
        this.stationSelected = "[Export All]";
        this.disableSelect = true;
    }

    async select_files() {
        try {
            fileDialog({multiple: true}).then(files => {
                // Create New Workbook
                this.mergeWorkbook = XLSX.utils.book_new();
                let prefix = xlUtil.common_prefix(files);
                // Parse all CSV files
                Promise.all(Array.prototype.map.call(files, (file) => file.text())).then((file_texts) => {
                    // For each file data
                    file_texts.forEach(((text, i) => {
                        // Read into workbook and transfer sheet
                        let tempWorkbook = XLSX.read(text, {type: "binary"});
                        XLSX.utils.book_append_sheet(this.mergeWorkbook, tempWorkbook.Sheets[tempWorkbook.SheetNames[0]], files[i].name.substring(prefix.length, 30 + prefix.length).split(".")[0]);
                    }));
                    // Populate
                    let sheet: XLSX.WorkSheet = this.mergeWorkbook.Sheets["station_audit"];
                    let idRange = xlUtil.header_search(sheet, XLSX.utils.decode_range(sheet["!ref"]), "station"); // Keyed by name rn
                    let idArray = xlUtil.range_to_array(sheet, idRange, 1);

                    let nameRange = xlUtil.header_search(sheet, XLSX.utils.decode_range(sheet["!ref"]), "station");
                    let nameArray = xlUtil.range_to_array(sheet, nameRange, 1);

                    this.stations = [{id: "[Export All]", name: "[Export All]"}];
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
                XLSX.writeFile(this.mergeWorkbook, this.stationSelected + ".xlsx");
            }
        } catch (error) {
            console.error(error);
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
                resolve(res && res.value && res.value.url ? res.value.url : '');
            })
        });
    }

    async loadfileBase64(file: File): Promise<string> {
        return new Promise<string>((resolve, reject) => {
            const reader = new FileReader();
            reader.readAsDataURL(file);
            reader.onload = () => resolve(reader.result.toString().replace(/^data:(.*,)?/, '') + '='.repeat(4 - (reader.result.toString().replace(/^data:(.*,)?/, '').length % 4)));
            reader.onerror = error => reject(error);
        })
    }

    async addBindingsHandler(callback: (event: Office.BindingDataChangedEventArgs | Office.BindingSelectionChangedEventArgs) => void, bindingName) {

    }

    async set_binding() {
        this.NameBinding.setDataAsync(this.field_stationName)
        this.VITBinding.addRowsAsync([["Cool", "Nice", "Sweet", "Good"]])
    }

    async initialize_binding() {
        try {
            Office.context.document.bindings.addFromNamedItemAsync("Bind_StationNameOpener", Office.BindingType.Text, {id: "Bind_StationName"}, (result => {
                if (result.status == Office.AsyncResultStatus.Failed) console.log(result.error.message)
                this.NameBinding = result.value
            }));
            Office.context.document.bindings.addFromNamedItemAsync("Bind_VIT", Office.BindingType.Table, {id: "Bind_VIT"}, (result => {
                if (result.status == Office.AsyncResultStatus.Failed) console.log(result.error.message)
                this.VITBinding = result.value as Office.TableBinding
            }));
        } catch (error) {
            console.error(error)
        }
    }

    async test_word() {
        try {
            await Word.run(async (context) => {
                let newDocument: DocumentCreated = context.application.createDocument(this.newDocumentBase64)
                newDocument.properties.customProperties.getItem("AF_StationName").value = this.stationSelected
                let nameParagraph = newDocument.body.insertParagraph(this.stationSelected + "_REV0", InsertLocation.start)
                newDocument.save()
                nameParagraph.delete()
                newDocument.open()
                await context.sync();
            });
        } catch (error) {
            console.error(error);
        }
    }
}

import XLSX from "xlsx"
import xlUtil from "./xlUtil"
import wordUtil from "./wordUtil"

import "reflect-metadata"
import { jsonObject, jsonMember, AnyT } from "typedjson"

@jsonObject
class SilverField {

  @jsonMember
  sheet_name: string
  @jsonMember
  field_name: string
  @jsonMember(() => SilverField.FieldOperation)
  operation: SilverField.FieldOperation
  @jsonMember(() => SilverField.FieldOutput)
  output: SilverField.FieldOutput

  dataRange: XLSX.Range
  outValue: Array<string> | string

  set_data(range: XLSX.Range, worksheet: XLSX.WorkSheet) {
    this.dataRange = range
    this.outValue = this.operation.output(worksheet, range)
  }
}

namespace SilverField {
  @jsonObject
  export class FieldOperation {

    @jsonMember
    operatorGetData: FieldOperation.OpGetData
    @jsonMember
    operatorCompare: FieldOperation.OpCompare
    @jsonMember
    operatorReduce: FieldOperation.OpReduce
    @jsonMember(AnyT)
    compareValue: string | number
    @jsonMember
    truthyValue: string
    @jsonMember
    falsyValue: string

    get_data(worksheet: XLSX.WorkSheet, range: XLSX.Range): Array<string | number> {
      console.log(FieldOperation.getDataFunctions.get(FieldOperation.OpGetData.string)(worksheet, range, 0))
      return FieldOperation.getDataFunctions.get(this.operatorGetData)(worksheet, range, 1)
    }

    compare(data: Array<string | number>): Array<boolean> {
      if (FieldOperation.compareFunctions.get(this.operatorCompare) != undefined) {
        return data.map((value) => FieldOperation.compareFunctions.get(this.operatorCompare)(value, this.compareValue))
      }
    }

    reduce(data: Array<string | number | boolean>): string | number | boolean {
      if (FieldOperation.reduceFunctions.get(this.operatorReduce) != undefined) {
        return data.reduce(FieldOperation.reduceFunctions.get(this.operatorReduce))
      }
    }

    stringify(data: Array<boolean | string | number> | boolean | string | number): string | Array<string> {
      if (Array.isArray(data)) {
        return data.map((value) => (typeof value == "boolean" ? (value ? this.truthyValue : this.falsyValue) : value.toString()))
      } else {
        return typeof data == "boolean" ? (data ? this.truthyValue : this.falsyValue) : data.toString()
      }
    }

    output(worksheet: XLSX.WorkSheet, range: XLSX.Range): Array<string> | string {
      let data = this.get_data(worksheet, range)
      if (data.length == 0) {
        return this.falsyValue
      }
      let compareResult = this.compare(data) ?? data
      return this.stringify(this.reduce(compareResult) ?? compareResult)
    }
  }

  export namespace FieldOperation {
    export enum OpGetData {
      string = "string",
      number = "number"
    }

    export const getDataFunctions: Map<OpGetData, (w: XLSX.WorkSheet, r: XLSX.Range, o: number) => Array<string | number>> = new Map([
      [OpGetData.string, ((w: XLSX.WorkSheet, r: XLSX.Range, o: number) => xlUtil.range_to_string_array(w, r, o) as Array<string>)],
      [OpGetData.number, ((w: XLSX.WorkSheet, r: XLSX.Range, o: number) => xlUtil.range_to_number_array(w, r, o) as Array<any>)] // Weird Behaviour
    ])

    export enum OpCompare {
      noop = "noop", // Don't generate boolean array
      numericLessThan = "numericLessThan",
      numericEquals = "numericEquals",
      numericGreaterThan = "numericGreaterThan",
      stringEquals = "stringEquals",
      stringNotEquals = "stringNotEquals"
    }

    export const compareFunctions: Map<OpCompare, (a: string | number, b: string | number) => boolean> = new Map([
      [OpCompare.numericLessThan, (a, b) => a < b],
      [OpCompare.numericEquals, (a, b) => a == b],
      [OpCompare.numericGreaterThan, (a, b) => a > b],
      [OpCompare.stringEquals, (a, b) => a.toString().toLowerCase() == b.toString().toLowerCase()],
      [OpCompare.stringNotEquals, (a, b) => a.toString().toLowerCase() != b.toString().toLowerCase()]
    ])

    export enum OpReduce {
      noop = "noop", // Leave as array
      count = "count",
      booleanCountIf = "booleanCountIf",
      numericSum = "numericSum",
      booleanAny = "booleanAny",
      booleanAll = "booleanAll"
    }

    export const reduceFunctions: Map<OpReduce, (acc: string | number | boolean, val: string | number | boolean) => string | number | boolean> = new Map([
      [OpReduce.count, (acc, val) => (acc as number) + 1],
      [OpReduce.booleanCountIf, (acc, val) => (acc as number) + (val ? 1 : 0)],
      [OpReduce.numericSum, (acc, val) => (acc as number) + (val as number)],
      [OpReduce.booleanAll, (acc, val) => (acc) && (val)],
      [OpReduce.booleanAny, (acc, val) => (acc) || (val)]
    ])
  }

  @jsonObject
  export class FieldOutput {

    @jsonMember
    method: FieldOutput.OutputMethod
    @jsonMember
    destinationName: string

    textBinding: Office.Binding

    bind() {
      switch (this.method) {
        case SilverField.FieldOutput.OutputMethod.documentProperty:
          return
        case SilverField.FieldOutput.OutputMethod.textField:
          try {
            Office.context.document.bindings.addFromNamedItemAsync(this.destinationName, Office.BindingType.Text, { id: this.destinationName }, (result => {
              if (result.status == Office.AsyncResultStatus.Failed) {
                console.log(result.error.message)
              } else {
                this.textBinding = result.value
              }
            }))
          } catch (error) {
            console.error(error)
          }
          return
        case SilverField.FieldOutput.OutputMethod.tableRow:
          return
      }
    }

    outputAsync(outValue: Array<string> | string): boolean {
      switch (this.method) {
        case SilverField.FieldOutput.OutputMethod.documentProperty:
          return false
        case SilverField.FieldOutput.OutputMethod.textField:
          this.textBinding.setDataAsync(outValue)
          return true
        case SilverField.FieldOutput.OutputMethod.tableRow:
          return false
      }
    }

    async outputWord(context: Word.RequestContext, outValue: Array<string> | string) {
      switch (this.method) {
        case SilverField.FieldOutput.OutputMethod.documentProperty:
          return
        case SilverField.FieldOutput.OutputMethod.textField:
          return
        case SilverField.FieldOutput.OutputMethod.tableRow:
          let tables = context.document.body.tables
          context.load(tables)
          tables.items.forEach((table) => {
            if (table.getCell(0, 0).value == this.destinationName) { // Absolutely rubbish method. Can't access table title
              table.addRows(Word.InsertLocation.end, 1, [Array.from(outValue.toString().repeat(table.rows.items[0].cellCount))])
            }
          })
          await context.sync().catch((e) => console.error(e))
          return
      }
    }
  }

  export namespace FieldOutput {
    export enum OutputMethod {
      documentProperty = "documentProperty",
      textField = "textField",
      tableRow = "tableRow",
    }
  }
}

export default SilverField
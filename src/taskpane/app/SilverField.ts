import XLSX from "xlsx"
import xlUtil from "./xlUtil"
import wordUtil from "./wordUtil"

import "reflect-metadata"
import { jsonObject, jsonMember, AnyT, TypedJSON, jsonArrayMember } from "typedjson"

@jsonObject
class SilverField {

  @jsonMember(() => SilverField.FieldInput)
  input: SilverField.FieldInput
  @jsonMember(() => SilverField.FieldTransformation)
  transformation: SilverField.FieldTransformation
  @jsonMember(() => SilverField.FieldOutput)
  output: SilverField.FieldOutput

  outValue: Array<string> | string

  set_data(worksheet: XLSX.WorkSheet, output_range: XLSX.Range, filter_ranges: Array<XLSX.Range>) {
    this.outValue = this.transformation.get_output_data(worksheet, output_range, filter_ranges)
  }
}

namespace SilverField {

  export function fromJSON(fields_json: any): Array<SilverField> {
    return Array.prototype.map.call(fields_json, (field) => TypedJSON.parse(field, SilverField))
  }

  export function extract_field_ranges(fields: Array<SilverField>, workbook: XLSX.WorkBook) {
    workbook.SheetNames.forEach((sheetName) => {
      let matchedFields = fields.filter((field) => field.input.sheet == sheetName)
      if (matchedFields) {
        let sheet = workbook.Sheets[sheetName]
        let search_result = xlUtil.header_search_multi(sheet, xlUtil.used_range(sheet), matchedFields.map((field) => [field.input.output_field].concat(field.input.filter_fields ?? [])).flat())
        matchedFields.forEach((field) => field.set_data(sheet, search_result.shift(), search_result.splice(0, field.input.filter_fields?.length ?? 0)))
      }
    })
  }

  @jsonObject
  export class FieldInput {
    @jsonMember
    sheet: string
    @jsonMember
    output_field: string
    @jsonArrayMember(String)
    filter_fields: Array<string>

  }

  @jsonObject
  export class FieldTransformation {

    // TODO: Clean this up - should be some kind of object that applies to both output and filters.
    // TODO: Really should be able to run both without any output fields *or* without any filter fields with a valid op

    @jsonMember
    opGetDataOutput: FieldTransformation.OpGetData
    @jsonMember
    opGetDataFilters: FieldTransformation.OpGetData
    @jsonMember
    opCompareOutput: FieldTransformation.OpCompare
    @jsonMember
    opCompareFilters: FieldTransformation.OpCompare
    @jsonMember
    opReduceOutput: FieldTransformation.OpReduce
    @jsonMember
    opReduceFilters: FieldTransformation.OpReduce
    @jsonMember(AnyT)
    valCompareOutput: string | number
    @jsonMember(AnyT)
    valCompareFilters: string | number
    @jsonMember
    valOutputTruthy: string
    @jsonMember
    valOutputFalsy: string

    get_output_data(worksheet: XLSX.WorkSheet, output_range: XLSX.Range, filter_ranges: Array<XLSX.Range>): Array<string> | string {
      // Bad.

      if (output_range == undefined){ // No Falsy Output - Column wasn't found.
        return
      }

      let output_data = FieldTransformation.get_data(worksheet, output_range, this.opGetDataOutput)
      if (output_data.length == 0) {
        return this.valOutputFalsy
      }
      if (filter_ranges?.length ?? 0 > 0) {
        let filter_data = filter_ranges.map((range) => FieldTransformation.compare(FieldTransformation.get_data(worksheet, range, this.opGetDataFilters), this.opCompareFilters, this.valCompareFilters))
        output_data = output_data.filter((_, i) => FieldTransformation.reduce(filter_data.map((column) => column[i]) ?? filter_data.map((column) => column[i]), this.opReduceFilters))
      }

      // Falsy if everything is filtered out
      if (output_data.length == 0) {
        return this.valOutputFalsy
      }

      let compareResult = FieldTransformation.compare(output_data, this.opCompareOutput, this.valCompareOutput) ?? output_data
      return FieldTransformation.stringify(FieldTransformation.reduce(compareResult, this.opReduceOutput) ?? compareResult, this.valOutputTruthy, this.valOutputFalsy)
    }
  }

  export namespace FieldTransformation {
    export enum OpGetData {
      string = "string",
      number = "number"
    }

    const getDataFunctions: Map<OpGetData, (w: XLSX.WorkSheet, r: XLSX.Range, o: number) => Array<string | number>> = new Map([
      [OpGetData.string, ((w: XLSX.WorkSheet, r: XLSX.Range, o: number) => xlUtil.range_to_string_array(w, r, o) as Array<any>)], // Weird Typing Behaviour
      [OpGetData.number, ((w: XLSX.WorkSheet, r: XLSX.Range, o: number) => xlUtil.range_to_number_array(w, r, o))]
    ])

    export function get_data(worksheet: XLSX.WorkSheet, range: XLSX.Range, op: OpGetData): Array<string | number> {
      return getDataFunctions.get(op)(worksheet, range, 1)
    }

    export enum OpCompare {
      // Can also be null - do not generate boolean array
      numericLessThan = "numericLessThan",
      numericEquals = "numericEquals",
      numericGreaterThan = "numericGreaterThan",
      stringEquals = "stringEquals",
      stringNotEquals = "stringNotEquals"
    }

    const compareFunctions: Map<OpCompare, (a: string | number, b: string | number) => boolean> = new Map([
      [OpCompare.numericLessThan, (a, b) => a < b],
      [OpCompare.numericEquals, (a, b) => a == b],
      [OpCompare.numericGreaterThan, (a, b) => a > b],
      [OpCompare.stringEquals, (a, b) => a.toString().toLowerCase() == b.toString().toLowerCase()],
      [OpCompare.stringNotEquals, (a, b) => a.toString().toLowerCase() != b.toString().toLowerCase()]
    ])

    export function compare(data: Array<string | number>, op: OpCompare, val: string | number): Array<boolean> {
      if (compareFunctions.get(op) != undefined) {
        return data.map((value) => compareFunctions.get(op)(value, val))
      }
    }

    export enum OpReduce {
      // Can also be null - do not aggregate
      count = "count",
      booleanCountIf = "booleanCountIf",
      numericSum = "numericSum",
      booleanAny = "booleanAny",
      booleanAll = "booleanAll",
      stringConcat = "stringConcat"
    }

    const reduceFunctions: Map<OpReduce, (acc: string | number | boolean, val: string | number | boolean) => string | number | boolean> = new Map([
      [OpReduce.count, (acc, _) => (acc as number) + 1],
      [OpReduce.booleanCountIf, (acc, val) => (acc as number) + (val ? 1 : 0)],
      [OpReduce.numericSum, (acc, val) => (acc as number) + (val as number)],
      [OpReduce.booleanAll, (acc, val) => (acc) && (val)],
      [OpReduce.booleanAny, (acc, val) => (acc) || (val)],
      [OpReduce.stringConcat, (acc, val) => `${acc}\n${val}`]
    ])

    export function reduce(data: Array<string | number | boolean>, op: OpReduce): string | number | boolean {
      if (reduceFunctions.get(op) != undefined) {
        return data.reduce(reduceFunctions.get(op))
      }
    }

    export function stringify(data: Array<boolean | string | number> | boolean | string | number, truthy: string, falsy: string): string | Array<string> {
      if (Array.isArray(data)) {
        return data.map((value) => (typeof value == "boolean" ? (value ? truthy : falsy) : value.toString()))
      } else {
        return typeof data == "boolean" ? (data ? truthy : falsy) : data.toString()
      }
    }

  }

  @jsonObject
  export class FieldOutput {

    @jsonMember
    method: FieldOutput.OutputMethod
    @jsonMember
    destinationName: string

    textBinding: Office.Binding

    bind(): Promise<boolean> {
      if (!this.textBinding) {
        switch (this.method) {
          case FieldOutput.OutputMethod.textField:
            return new Promise<boolean>(resolve => {
              wordUtil.bindTextPromise(this.destinationName).then((value) => {
                this.textBinding = value
                resolve(true)
              }).catch((error) => {
                console.error(error)
                resolve(false)
              })
            })
        }
        return Promise.resolve(null)
      } else {
        return Promise.resolve(true) // Already Bound
      }
    }

    outputAsync(outValue: Array<string> | string): boolean {
      switch (this.method) {
        case FieldOutput.OutputMethod.documentProperty:
          return null
        case FieldOutput.OutputMethod.textField:
          try {
            this.textBinding.setDataAsync(outValue ?? "xx")
            return true
          }
          catch (e) {
            console.error(e)
            return false
          }
        case FieldOutput.OutputMethod.card:
          return null
      }
      return null
    }

    async outputWord(context: Word.RequestContext, outValue: Array<string> | string): Promise<boolean> {
      try {
        switch (this.method) {
          case FieldOutput.OutputMethod.documentProperty:
            return Promise.resolve(null)
          case FieldOutput.OutputMethod.tableRow:
            let tables = context.document.body.tables
            context.load(tables)
            await context.sync()
            for (const table of tables.items) {
              let firstCell = table.getCell(0, 0)
              context.load(firstCell)
              await context.sync()
              if (firstCell.value == this.destinationName) {
                let tableRows = table.rows
                context.load(tableRows)
                await context.sync()
                //table.addRows(Word.InsertLocation.end, 1, [Array.from(outValue.toString().repeat(tableRows.items[table.rowCount - 1].cellCount))])
                table.addRows(Word.InsertLocation.end, 1, [[outValue.toString(), outValue.toString(), outValue.toString()]])
              }
            }
            await context.sync().catch((e) => console.error(e))
            return Promise.resolve(true)
        }
        return Promise.resolve(null)
      } catch (error) {
        console.error(error)
        return Promise.resolve(false)
      }
    }
  }

  export namespace FieldOutput {
    export enum OutputMethod {
      documentProperty = "documentProperty",
      textField = "textField",
      tableRow = "tableRow",
      card = "card"
    }
  }
}

export default SilverField
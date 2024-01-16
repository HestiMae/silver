import XLSX from "xlsx";
import xlUtil from "./xlUtil";

class SilverField {
  sheet_name: string
  field_name: string
  operation: SilverField.FieldOperation
  dataRange: XLSX.Range
  data: Array<string | number>
  dataBoolean: Array<boolean>


  constructor(sheet_name?: string, field_name?: string, operation?: SilverField.FieldOperation) {
    this.sheet_name = sheet_name
    this.field_name = field_name
    this.operation = Object.assign(new SilverField.FieldOperation(), operation)
    this.dataRange = undefined
  }

  set_data(range: XLSX.Range, worksheet: XLSX.WorkSheet)
  {
    this.dataRange = range;
    this.data = SilverField.FieldOperation.get_data(range, worksheet, this.operation.operatorCompare)
    //TODO: switch statement
    this.dataBoolean = this.operation.compare(this.data)
  }
}

namespace SilverField {
  export class FieldOperation {
    operatorCompare: FieldOperation.OpCompare
    operatorReduce: FieldOperation.OpCompare
    value: string | number
    constructor(operator?: FieldOperation.OpCompare, value?: string | number) {
      this.operatorCompare = operator
      this.value = value
    }
    compare(data: Array<string | number>) : Array<boolean>
    {
      if (this.operatorCompare != FieldOperation.OpCompare.noop) {
        return data.map((value) => FieldOperation.compareFunctions.get(this.operatorCompare)(value, this.value))
      }
    }
    reduce(data: Array<string | number | boolean>) {

    }
  }

  export namespace FieldOperation {
    export enum OpCompare {
      noop, // Don't generate boolean array
      numericLessThan,
      numericEquals,
      numericGreaterThan,
      stringEquals,
      stringNotEquals
    }
    export enum OpReduce {
      noop, // Leave as array
      count,
      booleanCountIf,
      numericSum
    }
    export const compareFunctions: Map<OpCompare, (a: string | number, b : string | number) => boolean> = new Map([
      [OpCompare.numericLessThan, (a, b) => a<b],
      [OpCompare.numericEquals, (a, b) => a==b],
      [OpCompare.numericGreaterThan, (a, b) => a>b],
      [OpCompare.stringEquals, (a, b) => a.toString().toLowerCase()==b.toString().toLowerCase()],
      [OpCompare.stringNotEquals, (a, b) => a.toString().toLowerCase() != b.toString().toLowerCase()]
    ])
    // export const reduceFunctions: Map<OpReduce, (a: Array<string | number | boolean>) => any> = new Map([
    //   [OpReduce.count, (a) => a.length()],
    //   [OpReduce.booleanCountIf, (a) => a],
    //   [OpReduce.numericSum, (a) => a],
    // ])


    export function get_data(range: XLSX.Range, worksheet: XLSX.WorkSheet, operator: OpCompare): Array<string | number>
    {
      if (operator == OpCompare.numericEquals
        || operator == OpCompare.numericLessThan
        || operator == OpCompare.numericGreaterThan)
      {
        return xlUtil.range_to_number_array(worksheet, this.dataRange, 1)
      }
      else if (operator == OpCompare.stringEquals || operator == OpCompare.stringNotEquals)
      {
        return xlUtil.range_to_string_array(worksheet, this.dataRange, 1)
      }
    }
  }
}

export default SilverField;
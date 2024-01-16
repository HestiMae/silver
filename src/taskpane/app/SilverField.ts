class SilverField {
  constructor(sheet_name: string, field_name: string, operation: SilverField.FieldOperation) {
  }

}

namespace SilverField {
  export class FieldOperation {
    constructor(operator: FieldOperation.FieldOperator, value: string | number) {

    }
  }

  export namespace FieldOperation {
    export enum FieldOperator {
      numericLessThan,
      numericEquals,
      numericGreaterThan,
      stringEquals,
      stringNotEquals
    }
  }
}

export default SilverField;
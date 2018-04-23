namespace BizPad {
    using System;
    using DevExpress.XtraRichEdit.Fields;
    using DevExpress.XtraRichEdit.Model;

    public class FieldCalculatorServiceEx : FieldCalculatorService, IFieldCalculatorService {
        protected override CalculatedFieldBase CreateInitializedCalculatedField(
            PieceTable pieceTable,
            Token firstToken,
            FieldScanner scanner) {

            Token chartToken = firstToken;
            if (chartToken != null) {
                if (String.Equals(chartToken.Value, "CHART", StringComparison.OrdinalIgnoreCase)) {
                    return CreateInitializedChartField(pieceTable,scanner);
                }
            }

            return base.CreateInitializedCalculatedField(
                pieceTable,
                firstToken,
                scanner);
            }

        static CalculatedFieldBase CreateInitializedChartField(PieceTable pieceTable, FieldScanner scanner) {
            CalculatedFieldBase field = new ChartField();
            if (field == null) {
                return null;
            }

            InstructionCollection instructions = ParseInstructions(scanner, field);
            field.Initialize(pieceTable, instructions);
            return field;
        }
    }
}

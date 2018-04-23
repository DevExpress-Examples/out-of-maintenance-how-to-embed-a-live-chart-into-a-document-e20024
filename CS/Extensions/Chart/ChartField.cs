namespace BizPad {
    using System.Collections.Generic;
    using DevExpress.XtraRichEdit.Fields;
    using DevExpress.XtraRichEdit.Model;
    using DevExpress.XtraRichEdit.Utils;

    public class ChartField : CalculatedFieldBase {
        public static readonly string FieldType = "CHART";
        protected override string FieldTypeName {
            get {
                return FieldType;
            }
        }

        static readonly Dictionary<string, bool> switchesWithArgument = CreateSwitchesWithArgument("w", "h", "d", "c", "s");
        protected override Dictionary<string, bool> SwitchesWithArguments {
            get {
                return switchesWithArgument;
            }
        }

        ChartImage chart = new ChartImage();

        public override void Initialize(InstructionCollection switches) {
            base.Initialize(switches);
            chart.Initialize(switches);
        }


        public override CalculatedFieldValue GetCalculatedValueCore(PieceTable sourcePieceTable,
            MailMergeDataMode mailMergeDataMode,
            Field documentField) {

            RichEditImage image = chart.CreateImage();
            DocumentModel targetModel = sourcePieceTable.DocumentModel.GetFieldResultModel();
            targetModel.MainPieceTable.InsertInlinePicture(DocumentLogPosition.Zero, image);
            return new CalculatedFieldValue(targetModel);
        }
    }
}

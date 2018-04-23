Imports DevExpress.Office.Utils
Imports DevExpress.XtraRichEdit.Model
Imports DevExpress.XtraRichEdit.Fields
Imports System.Collections.Generic

Namespace BizPad

    Public Class ChartField
        Inherits CalculatedFieldBase

        Public Shared ReadOnly FieldType As String = "CHART"
        Protected Overrides ReadOnly Property FieldTypeName() As String
            Get
                Return FieldType
            End Get
        End Property

        Private Shared ReadOnly switchesWithArgument As Dictionary(Of String, Boolean) = CreateSwitchesWithArgument("w", "h", "d", "c", "s")
        Protected Overrides ReadOnly Property SwitchesWithArguments() As Dictionary(Of String, Boolean)
            Get
                Return switchesWithArgument
            End Get
        End Property

        Private chart As New ChartImage()

        Public Overrides Sub Initialize(ByVal pieceTable As PieceTable, ByVal switches As InstructionCollection)
            MyBase.Initialize(pieceTable, switches)
            chart.Initialize(switches)
        End Sub


        Public Overrides Function GetCalculatedValueCore(ByVal sourcePieceTable As PieceTable, ByVal mailMergeDataMode As MailMergeDataMode, ByVal documentField As Field) As CalculatedFieldValue

            Dim image As OfficeImage = chart.CreateImage()
            Dim targetModel As DocumentModel = sourcePieceTable.DocumentModel.GetFieldResultModel()
            targetModel.MainPieceTable.InsertInlinePicture(DocumentLogPosition.Zero, image)
            Return New CalculatedFieldValue(targetModel)
        End Function
    End Class
End Namespace

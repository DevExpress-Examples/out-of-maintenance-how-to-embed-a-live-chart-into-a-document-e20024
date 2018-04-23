Imports Microsoft.VisualBasic
Imports System
	Imports DevExpress.XtraBars.Ribbon
Namespace BizPad

	Partial Public Class Editor
		Inherits RibbonForm
		Public Sub New()
			InitializeComponent()
			Me.richEdit.LoadDocument("Chart.docx")
		End Sub
	End Class
End Namespace

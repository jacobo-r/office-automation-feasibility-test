Option Explicit

' === CONFIGURATION (edit these two only) ===
Const BASE_PATH = "D:\USUARIOS (NO BORRAR)\ADMINISTRADOR\Downloads\office-automation-feasability-test-main"
Const PDF_EXPORT_DIR = "D:\USUARIOS (NO BORRAR)\ADMINISTRADOR\Documents\PDF_Exports"
' ===========================================

Private Function Prop(ByVal n$, ByVal v As Variant) As com.sun.star.beans.PropertyValue
    Dim p As New com.sun.star.beans.PropertyValue
    p.Name = n
    p.Value = v
    Prop = p
End Function

Private Function NowStamp$()
    Dim d As Date: d = Now
    NowStamp = Format(d, "yyyy-mm-dd_hh-nn-ss")
End Function

Private Function SafeTitle$(ByVal s$)
    Dim t$: t = s
    If t = "" Then t = "document"
    SafeTitle = Replace(t, " ", "_")
End Function

Sub ExportAndSendPDF()
    On Error GoTo Oops

    ' === Active document check ===
    Dim doc As Object
    doc = ThisComponent
    If IsNull(doc) Then
        MsgBox "No active document.", 48, "Export PDF"
        Exit Sub
    End If

    ' === Determine export filter based on document type ===
    Dim filterName$
    filterName = "writer_pdf_Export"
    If doc.supportsService("com.sun.star.sheet.SpreadsheetDocument") Then
        filterName = "calc_pdf_Export"
    End If

    ' === Build export file path ===
    Dim pdfPath$, pdfUrl$
    pdfPath = PDF_EXPORT_DIR & "\" & SafeTitle(doc.Title) & "_" & NowStamp() & ".pdf"
    pdfUrl = ConvertToURL(pdfPath)

    ' === Export to PDF ===
    Dim args(0) As New com.sun.star.beans.PropertyValue
    args(0) = Prop("FilterName", filterName)
    doc.storeToURL pdfUrl, args()

    ' === Build Python sender command ===
    Dim WS_URL$, senderScript$, cmd$
    WS_URL = "ws://127.0.0.1:9000"
    senderScript = BASE_PATH & "\ws_send_pdf.py"

    cmd = "python " & Quote(senderScript) & " " & Quote(WS_URL) & " " & Quote(pdfPath)
    Shell("cmd /c start ""sendpdf"" " & cmd, 0)

    MsgBox "Exported and sent: " & pdfPath, 64, "Done"
    Exit Sub

Oops:
    MsgBox "Error: " & Err & " - " & Error$, 16, "Export PDF"
End Sub

Private Function Quote$(ByVal s$)
    Quote = """" & s & """"
End Function

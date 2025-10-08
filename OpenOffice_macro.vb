Option Explicit

' === CONFIGURATION ===
Const BASE_PATH = "D:\USUARIOS (NO BORRAR)\ADMINISTRADOR\Downloads\office-automation-feasibility-test-main"
Const PDF_EXPORT_DIR = BASE_PATH & "\temp_pdf"
Const PYTHON_EXE = "C:\Users\Administrador\AppData\Local\Programs\Python\Python312\python.exe"
Const WS_URL = "ws://127.0.0.1:9000"
' ======================

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

Private Sub EnsureFolder(path$)
    On Error Resume Next
    MkDir path
    On Error GoTo 0
End Sub

Sub ExportAndSendPDF()
    On Error GoTo Oops

    Dim doc As Object
    doc = ThisComponent
    If IsNull(doc) Then
        MsgBox "No active document.", 48, "Export PDF"
        Exit Sub
    End If

    ' Ensure export folder exists
    Call EnsureFolder(PDF_EXPORT_DIR)

    ' Choose export filter
    Dim filterName$
    If doc.supportsService("com.sun.star.sheet.SpreadsheetDocument") Then
        filterName = "calc_pdf_Export"
    Else
        filterName = "writer_pdf_Export"
    End If

    ' Build PDF file path
    Dim pdfPath$, pdfUrl$
    pdfPath = PDF_EXPORT_DIR & "\" & SafeTitle(doc.Title) & "_" & NowStamp() & ".pdf"
    pdfUrl = ConvertToURL(pdfPath)

    Dim args(0) As New com.sun.star.beans.PropertyValue
    args(0) = Prop("FilterName", filterName)
    doc.storeToURL pdfUrl, args()

    ' === Build final command for CMD ===
    Dim cmd$
    cmd = "cmd /c ""cd /d """ & BASE_PATH & """ && """ & PYTHON_EXE & """ ws_send_pdf.py """ & WS_URL & """ """ & pdfPath & """"""

    ' === Debug: show command before running (optional)
    ' MsgBox "Running: " & cmd

    ' === Execute silently ===
    Shell cmd, 0

    MsgBox "Exported and sent: " & pdfPath, 64, "Done"
    Exit Sub

Oops:
    MsgBox "Error " & Err & ": " & Error$, 16, "Export PDF"
End Sub

Private Function Quote$(ByVal s$)
    Quote = """" & s & """"
End Function

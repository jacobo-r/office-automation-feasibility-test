Option Explicit

' === CONFIGURATION ===
Const BASE_PATH = "D:\USUARIOS (NO BORRAR)\ADMINISTRADOR\Downloads\office-automation-feasibility-test-main"
Const PDF_EXPORT_DIR = BASE_PATH & "\temp_pdf"
Const PYTHON_EXE = "C:\Users\Administrador\AppData\Local\Programs\Python\Python314\python.exe"
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

Private Function Quote$(ByVal s$)
    Quote = """" & s & """"
End Function

' === MAIN ROUTINE ===
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

    ' Build PDF path
    Dim pdfPath$, pdfUrl$
    pdfPath = PDF_EXPORT_DIR & "\" & SafeTitle(doc.Title) & "_" & NowStamp() & ".pdf"
    pdfUrl = ConvertToURL(pdfPath)

    Dim args(0) As New com.sun.star.beans.PropertyValue
    args(0) = Prop("FilterName", filterName)
    doc.storeToURL pdfUrl, args()

    ' === Build final command for CMD ===
    Dim cmd$, logPath$
    logPath = BASE_PATH & "\ws_send_log_from_macro.txt"

    cmd = "cmd /k ""cd /d """ & BASE_PATH & """ && """ & PYTHON_EXE & """ ws_send_pdf.py """ & WS_URL & """ """ & PDF_EXPORT_DIR & """ >> """ & logPath & """ 2>&1"""

    ' Optional: small wait to ensure file flush
    Wait 1000

    ' Execute visible (1 = show console)
    Shell cmd, 1

    MsgBox "Exported and sending: " & pdfPath, 64, "Done"
    Exit Sub

Oops:
    MsgBox "Error " & Err & ": " & Error$, 16, "Export PDF"
End Sub

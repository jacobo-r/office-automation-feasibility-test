Option Explicit

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

    Dim doc As Object
    doc = ThisComponent
    If IsNull(doc) Then
        MsgBox "No active document.", 48, "Export PDF"
        Exit Sub
    End If

    Dim filterName$
    filterName = "writer_pdf_Export"
    If doc.supportsService("com.sun.star.sheet.SpreadsheetDocument") Then
        filterName = "calc_pdf_Export"
    End If

    ' Export to Desktop
    Dim home$, pdfPath$, pdfUrl$
    home = Environ("HOME")
    pdfPath = home & "/Desktop/" & SafeTitle(doc.Title) & "_" & NowStamp() & ".pdf"
    pdfUrl = ConvertToURL(pdfPath)

    Dim args(0) As New com.sun.star.beans.PropertyValue
    args(0) = Prop("FilterName", filterName)
    doc.storeToURL pdfUrl, args()

    ' Call Python sender from your project directory
    Dim WS_URL$, cmd$
    WS_URL = "ws://127.0.0.1:9000"  ' local test server
    cmd = "python3 " & Quote("/Users/jacoboruiz/dev/office_auto_feasibility/ws_send_pdf.py") & _
          " " & Quote(WS_URL) & " " & Quote(pdfPath)
    Shell(cmd, 0)

    MsgBox "Exported and sent: " & pdfPath, 64, "Done"
    Exit Sub

Oops:
    MsgBox "Error: " & Err & " - " & Error$, 16, "Export PDF"
End Sub

Private Function Quote$(ByVal s$)
    Quote = """" & s & """"
End Function

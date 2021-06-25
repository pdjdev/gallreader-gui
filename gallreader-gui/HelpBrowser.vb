Public Class HelpBrowser
#Region "브라우저 확대/축소"
    'Code by Clive Dela Cruz (https://itsourcecode.com/free-projects/vb-net/zoom-webbrowser-using-vb-net/)

    Public Enum Exec
        OLECMDID_OPTICAL_ZOOM = 63
    End Enum

    Private Enum execOpt
        OLECMDEXECOPT_DODEFAULT = 0
        OLECMDEXECOPT_PROMPTUSER = 1
        OLECMDEXECOPT_DONTPROMPTUSER = 2
        OLECMDEXECOPT_SHOWHELP = 3
    End Enum

    Dim dpivalue As Double

    Public Sub PerformZoom(Browser As WebBrowser, Value As Integer)
        Try
            Dim Res As Object = Nothing
            Dim MyWeb As Object
            MyWeb = Browser.ActiveXInstance
            MyWeb.ExecWB(Exec.OLECMDID_OPTICAL_ZOOM, execOpt.OLECMDEXECOPT_PROMPTUSER, CObj(Value), CObj(IntPtr.Zero))
        Catch ex As Exception
            MsgBox(ex.Message)
        End Try
    End Sub

#End Region

    Private Sub HelpBrowser_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Dim g As Graphics = CreateGraphics()
        dpivalue = g.DpiX '기본 = 96
        WebBrowser1.Navigate("https://pdjdev.github.io/gallreader-gui/help/")
    End Sub

    Private Sub WebBrowser1_DocumentCompleted(sender As Object, e As WebBrowserDocumentCompletedEventArgs) Handles WebBrowser1.DocumentCompleted
        PerformZoom(WebBrowser1, Convert.ToInt32(dpivalue * dpivalue / 100))
    End Sub
End Class
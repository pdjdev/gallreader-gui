Module DataModule
    '중간의 문자열을 리턴하는 함수
    Public Function midReturn(ByVal first As String, ByVal last As String, ByVal total As String) As String
        If total.Contains(first) Then
            Dim FirstStart As Long = total.IndexOf(first) + first.Length + 1
            Return Trim(Mid$(total, FirstStart, total.Substring(FirstStart).IndexOf(last) + 1))
        Else
            Return Nothing
        End If
    End Function

    'xml형식 파일을 전체값에서 따로 추출하는 함수
    Public Function getData(datastr As String, name As String) As String
        Return midReturn("<" + name + ">", "</" + name + ">", datastr)
    End Function
End Module

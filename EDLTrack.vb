Public Class EDLTrack
    Public myoptions As New List(Of EDLOption)
    Public Function GetOption(name As String) As String
        Dim myop As List(Of EDLOption)
        myop = myoptions.Where(Function(x) x.name = name).ToList()
        If myop.Count > 0 Then
            Return myop(0).value
        Else
            Return Nothing
        End If
    End Function
    Public Sub Initialize(edlformat As String, values As String)
        Dim myvaluez As String()
        myvaluez = edlformat.Split(New String() {";"}, StringSplitOptions.RemoveEmptyEntries)
        Initialize(myvaluez, values)
    End Sub

    Public Sub Initialize(edlformat As String(), values As String)
        Dim myvaluezz As String()
        myvaluezz = values.Split(New String() {";"}, StringSplitOptions.RemoveEmptyEntries)
        For i = 1 To edlformat.Length
            myoptions.Add(New EDLOption(Resyntax(edlformat(i - 1)), Resyntax(myvaluezz(i - 1))))
        Next
    End Sub

    Public Sub New()
    End Sub

    Public Sub New(v1 As String, v2 As String)
        Initialize(v1, v2)
    End Sub

    Private Function Resyntax(v As String) As String
        Dim mystr As String
        mystr = v.Replace(";", "").Replace(ChrW(34), "")
        While mystr.StartsWith(" ")
            mystr = mystr.Substring(1, mystr.Length - 1)
        End While
        While mystr.EndsWith(" ")
            mystr = mystr.Substring(0, mystr.Length - 1)
        End While
        Return mystr
    End Function
End Class

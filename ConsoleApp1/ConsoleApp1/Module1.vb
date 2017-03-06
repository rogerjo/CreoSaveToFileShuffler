Module Module1

    Sub Main()

        Call snb

    End Sub

    Private Sub snb()
        Dim c00 As String
        c00 = "One,Two,""Three,Four"",Five"

        Dim sp As String()

        sp = Split(c00, Chr(34))
        For j = 0 To UBound(sp) Step 2
            sp(j) = Replace(sp(j), ",", "_")
        Next
        Dim sn As String()
        sn = Split(Join(sp, Chr(34)), "_")

        MsgBox(sn.ToString)
    End Sub
End Module

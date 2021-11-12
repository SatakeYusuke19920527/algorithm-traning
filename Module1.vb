Module Module1

    Sub Main()
        Dim s() As String = {"", ",", ",,", ",,,", ""}
        s = {"", "\", "\\", "\\\", ""}
        's = {"\,", "\\,", "\\\,"}
        's = {"aabb", "aa,bb", "aa\,bb", "", "aa,,bb", "aa\\bb"}
        s = {vbBack, vbNewLine, vbLf, vbCr, vbCrLf}

        Dim sr As String = ",\,,\,\,,\,\,\,,"
        sr = ",\\,\\\\,\\\\\\,"
        'sr = "\\\,,\\\\\,,\\\\\\\,"
        'sr = "aabb,aa\,bb,aa\\\,bb,,aa\,\,bb,aa\\\\bb"
        sr = String.Format("{0},{1},{2},{3},{4}", vbBack, vbNewLine, vbLf, vbCr, vbCrLf)

        Dim x As String = CombineByComma(s)

        ' test1
        If x <> sr Then
            MsgBox("test1 error")
            Return
        End If

        Dim r() As String = SplitByComma(x)

        ' test2
        If s.Length <> r.Length Then
            MsgBox("test2 error")
            Return
        End If

        ' test3
        For i As Integer = 0 To s.Length - 1
            If s(i) <> r(i) Then
                MsgBox("test2 error")
                Return
            End If
        Next

        MsgBox("success")

    End Sub

    Private Function CombineByComma(s() As String) As String
        '最終行かどうか判断する為のカウンタ
        Dim loopCount As Integer = 0
        '文字列連結用
        Dim sb As New System.Text.StringBuilder()

        For Each sElement As String In s
            '一文字づつ分割
            For i = 0 To sElement.Length - 1
                ',がきたら\,、\がきたら\\、それ以外ならそのまま追加
                If sElement.Chars(i) = "," Then
                    Dim str As String = String.Format("{0}{1}", "\", sElement.Chars(i))
                    sb.Append(str)
                ElseIf sElement.Chars(i) = "\" Then
                    Dim str As String = String.Format("{0}{1}", "\", sElement.Chars(i))
                    sb.Append(str)
                Else
                    Dim str As String = String.Format("{0}", sElement.Chars(i))
                    sb.Append(str)
                End If
            Next i

            'ループの最後はコンマをつけない
            loopCount += 1
            If s.Length = loopCount Then
                Exit For
            End If
            sb.Append(",")
        Next

        Dim returnStr As String = sb.ToString()
        Return returnStr
    End Function

    Private Function SplitByComma(x As String) As String()
        Console.WriteLine("in SplitByComma")
        Console.WriteLine(x)
        Dim a() As String = New String() {"test1", "test2"}
        Return a
    End Function
End Module

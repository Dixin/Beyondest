<%
' ====================
' Beyondest.Com v3.6.1
' http://beyondest.com
' ====================
Function jk_pagecute(maxpage,thepages,viewpage,pageurl,pp,font_color)
    Dim pn
    Dim pi
    Dim page_num
    Dim ppp
    Dim pl
    Dim pr:pi = 1
    ppp    = pp\2
    If pp Mod 2 = 0 Then ppp = ppp - 1
    pl     = viewpage - ppp
    pr     = pl + pp - 1

    If pl < 1 Then
        pr = pr - pl + 1:pl = 1
        If pr > thepages Then pr = thepages
    End If

    If pr > Int(thepages) Then
        pl = pl + thepages - pr:pr = thepages
        If pl < 1 Then pl = 1
    End If

    If pl > 1 Then
        jk_pagecute = jk_pagecute & " <a href='" & pageurl & "' title='第一页'>[|<]</a> " & _
        " <a href='" & pageurl & "page=" & pl - 1 & "' title='上一页'>[<]</a> "
    End If

    For pi = pl To pr

        If CInt(viewpage) = CInt(pi) Then
            jk_pagecute = jk_pagecute & " <font color=" & font_color & ">[" & pi & "]</font> "
        Else
            jk_pagecute = jk_pagecute & " <a href='" & pageurl & "page=" & pi & "' title='第 " & pi & " 页'>[" & pi & "]</a> "
        End If

    Next

    If pr < thepages Then
        jk_pagecute = jk_pagecute & " <a href='" & pageurl & "page=" & pi & "' title='后一页'>[>]</a> " & _
        " <a href='" & pageurl & "page=" & thepages & "' title='最后一页'>[>|]</a> "
    End If

End Function %>
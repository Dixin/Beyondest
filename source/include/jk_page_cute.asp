<%
' ====================
' Beyondest.Com v3.6.1
' http://beyondest.com
' ====================
Function pagecute_fun(viewpage,thepages,pagecuteurl)
    Dim re_color,pf0,pf1,pf2,pf3,pf4,pf5
    re_color         = "#c0c0c0"
    pf0              = "已是第一页"
    pf1              = "第一页"
    pf2              = "上一页"
    pf3              = "下一页"
    pf4              = "最后一页"
    pf5              = "已是最后一页"
    pagecute_fun     = VbCrLf & "<table border=0 cellspacing=0 cellpadding=0><tr><form action='" & pagecuteurl & "' method=post><td>"

    If CInt(viewpage) = 1 Then
        pagecute_fun = pagecute_fun & VbCrLf & "<font color=" & re_color & ">" & pf0 & "</font>&nbsp;"
    Else
        pagecute_fun = pagecute_fun & VbCrLf & "<a href='" & pagecuteurl & "page=1' alt='" & pf1 & "'>" & pf1 & "</a>┋<a href='" & pagecuteurl & "page=" & CInt(viewpage) - 1 & "' alt='" & pf2 & "'>" & pf2 & "</a>&nbsp;"
    End If

    If CInt(viewpage) = CInt(thepages) Then
        pagecute_fun = pagecute_fun & VbCrLf & "<font color=" & re_color & " alt='" & pf5 & "'>" & pf5 & "</font>"
    Else
        pagecute_fun = pagecute_fun & VbCrLf & "<a href='" & pagecuteurl & "page=" & CInt(viewpage) + 1 & "' alt='" & pf3 & "'>" & pf3 & "</a>┋<a href='" & pagecuteurl & "page=" & CInt(thepages) & "' alt='" & pf4 & "'>" & pf4 & "</a>"
    End If

    If CInt(thepages) <> 1 Then
        pagecute_fun = pagecute_fun & VbCrLf & "&nbsp;<input type=text name=page value='" & viewpage & "' size=2>&nbsp;<input type=submit value='GO'>"
    End If

    pagecute_fun = pagecute_fun & VbCrLf & "</td></form></tr></table>"
End Function %>
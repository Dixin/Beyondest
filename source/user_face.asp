<!-- #include file="INCLUDE/config_user.asp" -->
<!-- #include file="INCLUDE/jk_pagecute.asp" -->
<!-- #include file="include/conn.asp" -->
<%
' ====================
' Beyondest.Com v3.6.1
' http://beyondest.com
' ====================

Dim nummer,page,rssum,thepages,viewpage,pageurl
thepages = 0:viewpage = 1:pageurl = "?"
tit      = "浏览头像"

Call web_head(0,0,3,0,0)
'------------------------------------left----------------------------------
Response.Write ukong

Call user_face()

Response.Write kong
'---------------------------------center end-------------------------------
Call web_end(0)

Sub user_face()
    Dim fc,j,fnum,cnum,rnum,tt,nnum
    cnum = 5:rnum = 3:fc = 0
    fnum = Int(web_var(web_num,11)) + 1
    'fc=fnum\cnum
    'if fnum mod cnum >0 then fc=fc+1

    nummer = cnum*rnum
    rssum  = fnum
    Call format_pagecute()

    If Int(viewpage) > 1 Then
        fc = (viewpage - 1)*nummer
    End If %>
<table border=0 cellpadding=0 cellspacing=4 align=center>
<tr align=center>
<td>本站共有 <font class=red><% Response.Write rssum %></font> 个头像</td>
<td width=10></td>
<td>页次：<font class=red><% Response.Write viewpage %></font>/<font class=red><% Response.Write thepages %></font></td>
<td width=10></td>
<td>分页：<% Response.Write jk_pagecute(nummer,thepages,viewpage,pageurl,5,"#ff0000") %></td>
</tr>
</table>
<table border=0 width='98%' cellpadding=0 cellspacing=8 align=center>
<%

    For j = 1 To rnum
        Response.Write vbcrlf & "<tr align=center>"

        For i = 1 To cnum
            nnum = cnum*(j - 1) + i - 1 + fc
            If nnum >= rssum Then Exit For
            tt = "<table border=0><tr><td align=center><img src='images/face/" & nnum & ".gif' border=0></td></tr><tr><td align=center><b>" & nnum & "</b></td></tr></table>"
            Response.Write vbcrlf & "<td>" & format_k(tt,1,5,120,120) & "</td>"
        Next

        Response.Write vbcrlf & "</tr>"
    Next %>
</table>
<%
End Sub %>
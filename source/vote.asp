<!--#include file="include/config.asp"-->
<!-- #include file="include/conn.asp" -->
<%
' ====================
' Beyondest.Com v3.6.1
' http://beyondest.com
' ====================

Dim id,types,mcolor,bgcolor,counter,ttt,j,c,w,h:j = 0:c = 0
w     = 350:h = 220
id    = Trim(Request.querystring("id"))
If Not(IsNumeric(id)) Then id = 0
types = Trim(Request.querystring("types"))
If Not(IsNumeric(types)) Then types = 1

If types = 1 Then
    ttt = "radio"
Else
    ttt = "checkbox"
End If

mcolor  = code_form(Request.querystring("mcolor"))
If Len(mcolor) <> 6 Then mcolor = "CC3300"
bgcolor = "#" & code_form(Request.querystring("bgcolor"))
If Len(bgcolor) <> 7 Then bgcolor = web_var(web_color,6)

sql    = "select id,vname,counter from vote where vid=" & id & " order by id"
Set rs = conn.execute(sql)

If rs.eof And rs.bof Then
    rs.Close:Set rs = Nothing
    Call close_conn()
    Response.Write "document.write(""没有此调查列表！"");"
    Response.End
End If

Response.Write vbcrlf & "document.write(""<table border=0 cellspacing=0 cellpadding=2>"");"
Response.Write vbcrlf & "document.write(""<form action='votetype.asp?action=save&vid=" & id & "' method=POST target='vote_view'>"");"

Do While Not rs.eof

    If j = 0 Then
        Response.Write vbcrlf & "document.write(""<tr><td align=center height=25><font color=#" & mcolor & "><b>" & code_html(rs("vname"),1,0) & "</b></font></td></tr>"");"
    Else
        Response.Write vbcrlf & "document.write(""<tr><td><input type=" & ttt & " name=vote_id value='" & rs("id") & "' style='background-color:" & web_var(web_color,6) & "'>" & code_html(rs("vname"),1,0) & "</td></tr>"");"
    End If

    j = j + 1:c = c + rs("counter")
    rs.movenext
Loop

rs.Close:Set rs = Nothing
Response.Write vbcrlf & "document.write(""<tr><td align=center height=25><input onclick=\""javascript:open_win('','vote_view'," & w & "," & h & ",'no');\"" type=submit value='投票'>&nbsp;&nbsp;<a href='javascript:;' onclick=\""javascript:open_win('votetype.asp?action=view&vid=" & id & "','vote_view'," & w & "," & h & ",'no');\"">查看结果</a><font class=gray>(共<font class=blue>" & c & "</font>票)</font></td></tr>"");"
Response.Write vbcrlf & "document.write(""</form></table>"");"

Call close_conn() %>
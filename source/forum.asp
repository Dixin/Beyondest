<!-- #include file="INCLUDE/config_forum.asp" -->
<!-- #include file="include/conn.asp" -->
<%
' ====================
' Beyondest.Com v3.6.1
' http://beyondest.com
' ====================

tit = tit_fir
tit_fir = ""

forum_mode = "half"	'forum_mode设为"full"时为全屏显示,"half"

If Len(Trim(Request.querystring("mode"))) > 0 Then
    forum_mode = Trim(Request.querystring("mode"))
End If

If forum_mode = "full" Then
    Call mode_full()
Else
    Call mode_half()
End If

Sub mode_half()
    Call web_head(0,0,1,0,0)
    '-----------------------------------center---------------------------------
    Response.Write kong & format_img("rbbs.jpg") & kong & gang
    Response.Write ukong
    Call forum_main("fk4")
    Call forum_down(Int(Mid(web_setup,3,1)))
    '---------------------------------center end-------------------------------
    Call web_center(1)
    '------------------------------------right---------------------------------
    Call format_login() %>
<table border=0 width='100%' cellspacing=0 cellpadding=0 align=center>
<tr><td align=center><% Call main_stat("fk4","jt1",0,5,1) %></td></tr>
<tr><td height=10></td></tr>
<tr><td align=center><% Call forum_cast("fk4","jt0",10,11) %></td></tr>
<tr><td height=10></td></tr>
<tr><td align=center><% Call forum_new("fk4","jt0",0,15,11,1) %></td></tr>
<tr><td height=5></td></tr>
<tr><td align=center><% Response.Write left_action("jt13",1) %></td></tr>
<tr><td height=10></td></tr>
</table>
<%
    '----------------------------------right end-------------------------------
    Call web_end(0)
End Sub

Sub mode_full()
    Call web_head(0,0,2,0,0)
    '-----------------------------------center---------------------------------
    Response.Write ukong %>
<table border=0 width='98%' cellspacing=0 cellpadding=0>
<tr align=center valign=top>
<td width='24%'><% Call format_login() %></td>
<td width='24%'><% Call main_stat("fk4","jt1",0,5,1) %></td>
<td width='1%'></td>
<td width='25%'><% Call forum_new("fk4","jt0",0,5,13,1) %></td>
<td width='1%'></td>
<td width='25%'><% Call forum_cast("fk4","jt0",5,13) %></td>
</tr>
</table>
<%
    Call forum_main("fk4")
    Call forum_down(Int(Mid(web_setup,3,1)))
    '---------------------------------center end-------------------------------
    Call web_end(0)
End Sub %>
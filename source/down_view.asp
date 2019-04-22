<!-- #include file="include/config_down.asp" -->
<%
' ====================
'                     Beyondest.Com V4.6 Demo版
' 
' http://beyondest.com
' ====================

Dim id:id = Trim(Request.querystring("id"))

If Not(IsNumeric(id)) Then
    Call format_redirect("down.asp")
    reponse.End
End If %>
<!-- #include file="include/jk_ubb.asp" -->
<!-- #include file="include/config_review.asp" -->
<!-- #include file="include/conn.asp" -->
<%
Dim cname
Dim sname
Dim temp1
Dim keyes
Dim power
Dim userp
Dim emoney
Dim url1
Dim url2
Dim sql2
Dim rs2
Set rs = Server.CreateObject("ADODB.recordset")
sql    = "select * from down where hidden=1 and id=" & id
rs.open sql,conn,1,1

If rs.eof And rs.bof Then
    rs.Close:Set rs = Nothing
    Call close_conn()
    Call format_redirect("down.asp")
    Response.End
End If

cid    = rs("c_id")
sid    = rs("s_id")
keyes  = rs("keyes")
power  = rs("power")
emoney = rs("emoney")

cname  = "音乐浏览":sname = ""

If cid > 0 Then

    If sid > 0 Then
        sql2  = "select jk_class.c_name,jk_sort.s_name from jk_sort inner join jk_class on jk_sort.c_id=jk_class.c_id where jk_sort.c_id=" & cid & " and jk_sort.s_id=" & sid
    Else
        sql2  = "select c_name from jk_class where c_id=" & cid
    End If

    Set rs2   = conn.execute(sql2)

    If Not (rs2.eof And rs2.bof) Then
        cname = rs2("c_name"):tit = cname
        If sid > 0 Then sname = rs2("s_name"):tit = sname & "（" & cname & "）"
    End If

    rs2.Close:Set rs2 = Nothing
End If

If action = "download" Then
    Call web_head(1,0,0,0,0)
Else
    Call web_head(0,0,0,0,0)
End If

'--------------------------------download---------------------------------
userp = Int(format_power(login_mode,2))

If action = "download" Then
    Call emoney_notes(power,emoney,n_sort,id,"js",1,1,"?id=" & id)

    If Trim(Request.querystring("url")) = "download2" Then
        index_url = rs("url2")
    Else
        index_url = rs("url")
    End If

    rs.Close:Set rs = Nothing
    sql = "update down set counter=counter+1 where id=" & id
    conn.execute(sql)
    Call close_conn()

    Response.redirect "" & url_true(web_var(web_down,5),index_url) & ""
    Response.End
End If

'------------------------------------left---------------------------------- %>
<table border=0 width='96%' cellspacing=0 cellpadding=0 align=center>
<tr><td align=center><% Call format_login() %></td></tr>
<tr><td align=center><% Call down_sea() %></td></tr>
<tr><td align=center><% Call down_new_hot("jt0","","","","good",10,0,13,1,0) %></td></tr>
<tr><td align=center><% Call down_new_hot("jt0","","","","hot",10,0,13,1,0) %></td></tr>
<tr><td align=center><% Call down_new_hot("jt0","","","","new",10,0,13,1,0) %></td></tr>
</table>
<%
'----------------------------------left end--------------------------------
Call web_center(0)
'-----------------------------------center--------------------------------- %>
<table border=0 width='100%' cellspacing=0 cellpadding=0 align=center>
<tr><td width=1 bgcolor="<% Response.Write web_var(web_color,3) %>"></td>
<td align=center><% Call down_intro(sid,sname) %></td></tr>
<tr><td width=1 bgcolor="<% Response.Write web_var(web_color,3) %>"></td><td align=center>
  <table border=1 cellspacing=0 cellpadding=4 width='98%' bordercolorlight=<% Response.Write web_var(web_color,3) %> bordercolordark=<% Response.Write web_var(web_color,5) %>>
  <tr bgcolor=<% Response.Write web_var(web_color,6) %> bordercolordark=<% Response.Write web_var(web_color,5) %>>
  <td align=center colspan=3 height=30><font size=3 class=blue><b><% Response.Write rs("name") %></b></font></td></tr>
  <tr><td align=center width='15%' bgcolor=<% Response.Write web_var(web_color,5) %>>专辑类型：</td><td width='40%'><% Response.Write rs("genre") %>&nbsp;</td>
  <td align=center width='45%' rowspan=8>
   <%
Response.Write "<img src='images/down/" & rs("pic") & "' border=0>" %>
</td></tr>
  <tr><td align=center bgcolor=<% Response.Write web_var(web_color,5) %>>播放软件：</td><td><%
If rs("os") = "Realone" Then Response.Write "<a href=" & web_var(web_down,5) & "/soft/realoneplayer.rar><img src=images/down/tool_realone.gif alt='Real One Player'  border='0'></a>"
If rs("os") = "WinMediaPlayer" Then Response.Write "<a href=" & web_var(web_down,5) & "/soft/wmp2k.rar><img src=images/down/TOOL_WMP.gif alt='Windows Media Player for 98 & Me & 2k'  border='0'></a>&nbsp;&nbsp;&nbsp;&nbsp;<a href=" & web_var(web_down,5) & "/soft/wmpxp.rar><img src=images/down/TOOL_WMP.gif alt='Windows Media Player for XP'  border='0'></a>"
If rs("os") = "Winamp" Then Response.Write "<a href=" & web_var(web_down,5) & "/soft/Winamp.rar><img src=images/down/tool_winamp.gif alt='Winamp'  border='0'></a>" %>&nbsp;</td></tr>
  <tr><td align=center bgcolor=<% Response.Write web_var(web_color,5) %>>文件大小：</td><td><% Response.Write rs("sizes") %></td></tr>
  <tr><td align=center bgcolor=<% Response.Write web_var(web_color,5) %>>推荐等级：</td><td><img src='images/down/star<% Response.Write rs("types") %>.gif' border=0></td></tr>
  <tr><td align=center bgcolor=<% Response.Write web_var(web_color,5) %>>下载次数：</td><td><font class=red><% Response.Write rs("counter") %></font></td></tr>
  <tr><td align=center bgcolor=<% Response.Write web_var(web_color,5) %>>发&nbsp;布&nbsp;人：</td><td><% Response.Write format_user_view(rs("username"),1,1) %></td></tr>
  <tr><td align=center bgcolor=<% Response.Write web_var(web_color,5) %>>上传日期：</td><td><% Response.Write time_type(rs("tim"),88) %></td></tr>
  <tr><td align=center bgcolor=<% Response.Write web_var(web_color,5) %>>文件来自：</td><td><%
temp1 = rs("homepage")

If temp1 = "" Or IsNull(temp1) Or temp1 = "http://" Then
    Response.Write "<a href='" & web_var(web_config,2) & "' target=_blank>" & web_var(web_config,2) & "</a>"
Else
    Response.Write "<a href='" & temp1 & "' target=_blank>" & temp1 & "</a>"
End If %></td></tr>
  <tr><td align=center bgcolor=<% Response.Write web_var(web_color,5) %>>下载权限：</td><td colspan=2>&nbsp;注册用户</td></tr>
  <tr><td align=center bgcolor=<% Response.Write web_var(web_color,5) %>>下载地址：</td><td colspan=2>&nbsp;&nbsp;&nbsp;<a href='?action=download&id=<% Response.Write id %>'<% Response.Write atb %>><img src='IMAGES/DOWN/DOWNLOAD.GIF' border=0></a>&nbsp;
<% If Len(rs("url2")) > 8 Then %>
&nbsp;&nbsp;&nbsp;<a href='?action=download&url=download2&id=<% Response.Write id %>'<% Response.Write atb %>><img src='IMAGES/DOWN/download2.gif' border=0></a>
<% End If %></td></tr>
  <tr height=50 valign=top><td align=center bgcolor=<% Response.Write web_var(web_color,5) %>>作品备注：</td><td colspan=2><table borer=0 width='100%' class=tf><tr><td><%

temp1 = rs("remark")

If Len(temp1) < 3 Then
temp1 = "<font class=gray>好像没有关于该音乐的介绍哦！</font>"
Else
temp1 = code_jk(temp1)
End If

Response.Write temp1
rs.Close %></td></tr></table></td></tr>
  <tr valign=top><td align=center bgcolor=<% Response.Write web_var(web_color
Dim 5) %>>相关音乐：</td><td colspan=2><table border=0><% Dim tempsn
Dim tempcn
Dim sqls
Dim sqlt
Dim rss
Dim rst
sql    = "select id,name,tim,counter,c_id,s_id from down where hidden=1 and keyes like '%" & keyes & "%' and id<>" & id & " order by counter desc"
Set rs = conn.execute(sql)

If rs.eof And rs.bof Then
Response.Write vbcrlf & "<tr><td class=gray>没有与之相关的作品</td></tr>"
Else

Do While Not rs.eof
    temp1 = rs("name")
    sqls = "select s_name from jk_sort where s_id=" & rs("s_id")
    Set rss = conn.execute(sqls)
    tempsn = rss("s_name")
    rss.Close:Set rss = Nothing
    sqlt = "select c_name from jk_class where c_id=" & rs("c_id")
    Set rst = conn.execute(sqlt)
    tempcn = rst("c_name")
    rst.Close:Set rst = Nothing
    Response.Write vbcrlf & "<tr><td><img src=images/small/jt0.gif>&nbsp;" & tempsn & "（" & tempcn & "）：<a href='down_view.asp?id=" & rs("id") & "' title='" & code_html(temp1,1,0) & "'>" & code_html(temp1,1,30) & "</a></td></tr>"
    rs.movenext
Loop

End If

rs.Close:Set rs = Nothing %></table></td></tr>
  </table>
</td></tr>

<tr><td width=1 bgcolor="<% Response.Write web_var(web_color,3) %>"></td><td height=10></td></tr>
<tr><td width=1 bgcolor="<% Response.Write web_var(web_color,3) %>"></td><td align=center><% Call review_type(n_sort,id,"down_view.asp?id=" & id,1) %></td></tr>
<tr><td width=1 bgcolor="<% Response.Write web_var(web_color,3) %>"></td><td height=5></td></tr>
<tr><td width=1 bgcolor="<% Response.Write web_var(web_color,3) %>"></td><td align=center><% Call down_class_sortt(cid,sid) %></td></tr>
<tr><td width=1 bgcolor="<% Response.Write web_var(web_color,3) %>"></td><td align=center><% Call down_remark("jt0") %></td></tr>

</table>
<%
'---------------------------------center end-------------------------------
Call web_end(0) %>
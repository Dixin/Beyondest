<!-- #include file="config_counter.asp" -->
<%
' ====================
' Beyondest.Com v3.6.1
' http://beyondest.com
' ====================

Dim online_num,counter_s
online_num = 0:counter_s = ""
lefter     = VbCrLf & "<table border=0 width=777 cellspacing=0 cellpadding=0><tr><td width=1 bgcolor='" & web_var(web_color,3) & "'></td><td width=775 align=center >"
righter    = VbCrLf & "</td><td width=1 bgcolor='" & web_var(web_color,3) & "'></td></tr></table>"

Sub format_login() %>
<table border=0 width=185 cellpadding=0 cellspacing=0>
<tr><td align=center height=120 background='images/<% Response.Write web_var(web_config,5) %>/login_bg.gif'>
<%

    If login_username <> "" And login_mode <> "" Then %>
  <table border=0 cellspacing=0 cellpadding=0>
  <tr height=20><td></td></tr>
  <tr height=30><td align=center>你好，<b><font class=blue><% Response.Write login_username %></font></b></td></tr>
  <tr><td align=center>你现在已登陆 <font class=red><% Response.Write format_power(login_mode,1) %></font> 模式</td></tr>
  <tr height=30><td align=center><%
        Dim mess_dim

        If login_message > 0 Then
            mess_dim = "red"
            Response.Write "<bgsound src='images/mail/message.wav' border=0>"
        Else
            login_message = 0:mess_dim = "gray"
        End If

        Response.Write "<a href='user_mail.asp'><img src='images/mail/new.gif' align=absmiddle border=0>&nbsp;我的短信箱（<font class=" & mess_dim & ">" & login_message & "&nbsp;新</font>）</a>" %></td></tr>
  <tr><td align=center><a href='login.asp?action=logout'>退出登陆</a>&nbsp;┋&nbsp;<a href='user_main.asp'>用户中心</a></td></tr>
  </table>
<%
    Else %>
  <script language=javascript src='style/login.js'></script>
  <table border=0 cellspacing=0 cellpadding=0>
  <form name=login_frm method=post action='login.asp?action=login_chk' onsubmit="return login_true()">
  <input type=hidden name=re_log value='yes'>
  <tr height=16><td colspan=2></td></tr>
  <tr height=30><td>用户名称：</td><td><input type=text name=username size=14 maxlength=20></td></tr>
  <tr height=26><td>登陆密码：</td><td><input type=password name=password size=14 maxlength=20></td></tr>
  <tr height=26><td colspan=2 align=center>
    <table border=0 cellspacing=0 cellpadding=0><tr align=center valign=bottom>
    <td><a href='login.asp?action=register'>免费注册</a></td>
    <td width=60><a href='login.asp?action=nopass'>忘记密码</a></td>
    <td><input type=submit value='登 录'></td>
    </tr></table>
  </td></tr></form></table>
<% End If %>
</td></tr><tr><td height=5></td></tr></table>
<%

End Sub

Sub web_head(var1,var2,var3,var4,var5)
    Dim ttt,ntt,user_ip,user_sys,rs,sql,wt
    user_ip = ip_sys(1,1):user_sys = ip_sys(3,0)
    wt      = Int(Mid(web_setup,3,1))

    If web_login = 1 And index_url <> "error" Then

        If symbol_name(login_username) <> "yes" Or symbol_ok(login_password) <> "yes" Then
            login_mode                                    = "":login_username = "":login_password = ""
        Else
            sql                                           = "select id,power,popedom,emoney from user_data where hidden=1 and username='" & login_username & "' and password='" & login_password & "'"
            Set rs                                        = conn.execute(sql)

            If rs.eof And rs.bof Then
                Response.Cookies(web_cookies)("login_username") = ""
                Response.Cookies(web_cookies)("login_password") = ""
            Else
                login_popedom                             = rs("popedom"):login_mode = rs("power"):login_emoney = rs("emoney")
            End If

            rs.Close
        End If

        If wt = 1 And login_mode = "" Then
            ttt                                           = Request.cookies(web_cookies)("guest_name")
            Set rs                                        = conn.execute("select l_id from user_login where l_username='" & ttt & "'")

            If rs.eof And rs.bof Then
                ntt                                       = "游客" & Session.SessionID
                conn.execute("insert into user_login(l_username,l_type,l_where,l_tim_login,l_tim_end,l_ip,l_sys) values('" & ntt & "',1,'" & tit & "','" & now_time & "','" & now_time & "','" & user_ip & "','" & user_sys & "')")
                Response.cookies(web_cookies)("guest_name") = ntt
            Else
                conn.execute("update user_login set l_where='" & tit & "',l_tim_end='" & now_time & "' where l_id=" & rs("l_id"))
            End If

            rs.Close
        End If

        If (wt = 1 Or wt = 2) And login_mode <> "" Then
            login_message = 0
            sql           = "select count(*) from user_mail where accept_u='" & login_username & "' and types=1 and isread=0"
            Set rs        = conn.execute(sql)
            If Not(rs.eof And rs.bof) Then login_message = Int(rs(0))
            rs.Close
            sql    = "select l_id from user_login where l_username='" & login_username & "'"
            Set rs = conn.execute(sql)

            If rs.eof And rs.bof Then
                conn.execute("insert into user_login(l_username,l_type,l_where,l_tim_login,l_tim_end,l_ip,l_sys) values('" & login_username & "',0,'" & tit & "','" & now_time & "','" & now_time & "','" & user_ip & "','" & user_sys & "')")
            Else
                conn.execute("update user_login set l_where='" & tit & "',l_tim_end='" & now_time & "' where l_username='" & login_username & "'")
            End If

            rs.Close
            Response.cookies(web_cookies)("guest_name") = ""
        End If

        Set rs = conn.execute("select count(l_id) from user_login")
        If Not(rs.eof And rs.bof) Then online_num = rs(0)
        rs.Close:Set rs = Nothing
        If wt <> 1 And login_mode = "" Then online_num = online_num + 1
        counter_s = counter_type("view","no")
        '---------------------------------Access 2000----------------------------------
        'sql="delete from user_login where DateDiff('n',l_tim_end,now())>"&int(web_var(web_num,13))
        'sql="delete from user_login where l_tim_end<"&DateAdd("n",-"&int(web_var(web_num,13))&",now())
        'conn.execute(sql)
        Call cookies_yes()
    End If

    If page_power <> "" And format_page_power(login_mode) <> "yes" Then Call close_conn():Call cookies_type("power")

    Select Case var1
        Case 1
            If Int(web_var_num(web_setup,1,1)) <> 0 And (login_username = "" Or login_mode = "") Then Call close_conn():Call cookies_type("login")
        Case 2
            If login_username = "" Or login_mode = "" Then Call close_conn():Call cookies_type("login")
    End Select

    Select Case var2
        Case 1
            Call close_conn()
        Case 2

            Exit Sub
        End Select %>
<html>
<head>
<title><% Response.Write web_var(web_config,1) & " - " & tit %></title>
<meta name="Description"  content="Beyondest">
<meta name="keywords" content="最全的Beyond资料,最好的Beyond网站,asp,Beyondest,笼民">
<meta name="author" content="Beyondest">
<meta http-equiv=Content-Type content=text/html; charset=gb2312>
<link rel=stylesheet href='include/beyondest.css' type=text/css>
<script language=javascript src='style/open_win.js'></script>
<script language=javascript src='style/mouse_on_title.js'></script>
</head>
<body topmargin=0 leftmargin=0 bgcolor=<%
        Response.Write web_var(web_color,1)
        ttt = web_var(web_config,7)

        If ttt <> "" Then
            Response.Write " background='images/" & web_var(web_config,5) & "/" & ttt & ".gif'"
        End If %>><a name='top'></a><center>
<% If Int(var3) = 4 Then Exit Sub %>
<% Response.Write lefter %>
<table border=0 width='100%' cellspacing=0 cellpadding=0>
<tr><td height=5 bgcolor=<% Response.Write web_var(web_color,3) %>></td></tr>
<tr><td height=2 bgcolor=<% Response.Write web_var(web_color,4) %>></td></tr>
<tr><td><% = kong %></td></tr>
<tr><td align=center>
  <table border=0 width='100%' cellspacing=0 cellpadding=0 class=tf>
  <tr align=center height=80>
  <td width='33%'><a href='<% = web_var(web_config,2) %>' target=_blank><img src='images/<% = web_var(web_config,5) %>/top_logo.gif' border=0></a></td>
  <td ><object classid="clsid:D27CDB6E-AE6D-11cf-96B8-444553540000" codebase="http://download.macromedia.com/pub/shockwave/cabs/flash/swflash.cab#version=6,0,29,0" width="468" height="60">
    <param name="movie" value="images/<% = web_var(web_config,5) %>/banner.swf">
    <param name="quality" value="high">
    <embed src="images/<% = web_var(web_config,1) %>/banner.swf" quality="high" pluginspage="http://www.macromedia.com/go/getflashplayer" type="application/x-shockwave-flash" width="468" height="60"></embed>
  </object></td>
<td align=right width='7%'>
    <table border=0 cellspacing=0 cellpadding=1 align=center>
    <tr height=20><td><a class=top href="javascript:window.external.AddFavorite('<% Response.Write web_var(web_config,2) %>','<% Response.Write web_var(web_config,1) %>')" style='target: ' _self?>加入收藏</a></td></tr>
    <tr height=20><td><a class=top href='help.asp'>帮助中心</a></td></tr>
    <tr height=20><td><a class=top href='mailto:dixinyan@live.com'>联系我们</a></td></tr>
    
    </table>
  </td>  

  </tr></table>
</td></tr>
<tr><td valign=middle>

  <table border=0 cellspacing=0 cellpadding=0 class=tf align=center background='images/<% = web_var(web_config,5) %>/menu_bg.gif'>
  <%
            Response.Write "<tr align=center height=35>"
            Dim wdim
            wdim = Split(web_menu,"|")

            For i = 0 To UBound(wdim)
                Response.Write vbcrlf & "<td><a href='" & Left(wdim(i),InStr(wdim(i),":") - 1) & ".asp'>" & format_icon(Left(wdim(i),InStr(wdim(i),":") - 1)) & "</a>" & Right(wdim(i),Len(wdim(i)) - InStr(wdim(i),":")) & "</td>"
            Next

            Erase wdim
            Response.Write "</tr>" %></table>
</td></tr>
<tr><td height=1 bgcolor=<% Response.Write web_var(web_color,3) %>></td></tr>
</table>
<table border=0 cellpadding=0 cellspacing=0 width='100%' align=center>
<tr><td align=left height=25 width=620>&nbsp;&nbsp;
今天是：<% Response.Write FormatDateTime(now_time,1) & "&nbsp;" & WeekdayName(Weekday(now_time)) %>
&nbsp;&nbsp;您现在位于：&nbsp;<a href='./'><% Response.Write web_var(web_config,1) %></a>&nbsp;→&nbsp;
<%

            If tit_fir = "" Then
                Response.Write tit
            Else
                Response.Write "<a href='" & index_url & ".asp'>" & tit_fir & "</a>&nbsp;→&nbsp;" & tit
            End If

            Response.Write "</td><td width=1 background='images/" & web_var(web_config,5) & "/bg_dian2.gif'></td><td width=154 align=right bgcolor=" & web_var(web_color,1) & ">"

            If login_message > 0 Then
                Response.Write "<a href='user_mail.asp'><img src='images/mail/new.gif' align=absmiddle border=0>&nbsp;我的短信箱（<font class=red>" & login_message & "&nbsp;新</font>）</a>"
            Else %>
<marquee scrolldelay=120 scrollamount=4 onMouseOut="if (document.all!=null){this.start()}" onMouseOver="if (document.all!=null){this.stop()}"><script src='style/head_scroll.js'></script></marquee>
<%
            End If

            Response.Write "</td></tr></table>" & _
            vbcrlf & "<table border=0 cellpadding=0 cellspacing=0 width='100%' align=center><tr><td background='images/" & web_var(web_config,5) & "/bg_dian.gif'></td></tr></table>" & righter & web_left(var3)
        End Sub

        Function web_left(lt)

            Select Case lt
                Case 1
                    web_left = vbcrlf & "<table border=0 width=775 cellspacing=0 cellpadding=0><tr valign=top align=center><td height=300 width=580>"
                Case 2
                    web_left = vbcrlf & "<table border=0 width=775 cellspacing=0 cellpadding=0><tr><td align=center>"
                Case 3
                    web_left = vbcrlf & "<table border=0 width=775 cellspacing=0 cellpadding=0><tr><td height=300 align=center>"
                Case Else
                    web_left = vbcrlf & "<table border=0 width=775 cellspacing=0 cellpadding=0><tr valign=top align=center><td height=300 width=185 bgcolor=" & web_var(web_color,6) & ">"
            End Select

            web_left         = lefter & web_left
        End Function

        Sub web_center(ct)
            Dim ttt,ttl,ttr:ttt = "bg3"
            If ttt <> "" Then ttr = " background='images/" & web_var(web_config,5) & "/" & ttt & "r.gif'":ttl = " background='images/" & web_var(web_config,5) & "/" & ttt & "l.gif'"

            Select Case ct
                Case 1
                    Response.Write vbcrlf & "</td>" & _
                    vbcrlf & "<td width=1 bgcolor=" & web_var(web_color,14) & "></td>" & _
                    vbcrlf & "<td width=8 bgcolor=" & web_var(web_color,5) & ttr & "></td>" & _
                    vbcrlf & "<td width=1 bgcolor=" & web_var(web_color,14) & "></td>" & _
                    vbcrlf & "<td wdith=185 bgcolor=" & web_var(web_color,6) & ">"
                Case Else
                    Response.Write vbcrlf & "</td>" & _
                    vbcrlf & "<td width=1 bgcolor=" & web_var(web_color,14) & "></td>" & _
                    vbcrlf & "<td width=8 bgcolor=" & web_var(web_color,5) & ttl & "></td>" & _
                    vbcrlf & "<td width=1 bgcolor=" & web_var(web_color,14) & "></td>" & _
                    vbcrlf & "<td wdith=580>"
            End Select

        End Sub

        Function web_right()
            web_right = vbcrlf & "</td></tr></table>" & righter
        End Function

        Sub web_end(wt)
            If IsObject(rs) Then Set rs = Nothing
            If wt = 0 Then Call close_conn()
            Response.Write web_right() & lefter %>
<table border=0 width='100%' cellspacing=0 cellpadding=0>
<tr><td height=1 bgcolor=<% Response.Write web_var(web_color,3) %>></td></tr>
<tr><td class=end align=center height=20 bgcolor=<% Response.Write web_var(web_color,2) %>>
<a class=menu href='help.asp?action=about'>关于我们</a>&nbsp;┋
<a class=menu href='gbook.asp'>网站留言</a>&nbsp;┋
<a href='links.asp' class=menu>友情链接</a>&nbsp;┋
<a class=menu href='online.asp'>与我在线</a>&nbsp;┋
版本：<a href='help.asp' target=_blank class=menu><% Response.Write web_edition %></a>&nbsp;┋
<a class=menu href='admin_login.asp' target=_blank>管理</a>&nbsp;┋
<a href='#top' class=menu>TOP</a>
</td></tr>
<tr><td height=1 bgcolor=<% Response.Write web_var(web_color,3) %>></td></tr>
<tr><td align=center height=20><% Response.Write counter_s %></td></tr>
<tr><td align=center><% Response.Write web_var(web_config,1) & "&nbsp;<font class=gray>" & web_var(web_stamp,Int(Mid(web_setup,3,1)) + 1) & "</font>" %>&nbsp;┋页面执行时间：<font class=red_3><% Response.Write FormatNumber((timer() - timer_start)*1000,3) %></font> 毫秒</td></tr>
<tr><td align=center height=20><% Response.Write web_var(web_error,4) %></td></tr>
<tr><td height=2 bgcolor=<% Response.Write web_var(web_color,4) %>></td></tr>
<tr><td height=5 bgcolor=<% Response.Write web_var(web_color,3) %>></td></tr>
</table><% Response.Write righter %>
</center></body></html>
<%
        End Sub

        Sub web_end2(wt)
            If IsObject(rs) Then Set rs = Nothing
            If wt = 0 Then Call close_conn()
            Response.Write "</center></body></html>"
        End Sub %>
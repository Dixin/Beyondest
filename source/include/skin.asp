<!-- #include file="config_counter.asp" -->
<%
'*******************************************************************
'
'                     Beyondest.Com V3.6 Demo版
' 
'           网址：http://www.beyondest.com
' 
'*******************************************************************

dim online_num,counter_s
online_num=0:counter_s=""
lefter=VbCrLf & "<table border=0 width=777 cellspacing=0 cellpadding=0><tr><td width=1 bgcolor='"&web_var(web_color,3)&"'></td><td width=775 align=center >"
righter=VbCrLf & "</td><td width=1 bgcolor='"&web_var(web_color,3)&"'></td></tr></table>"
sub format_login()
%>
<table border=0 width=185 cellpadding=0 cellspacing=0>
<tr><td align=center height=120 background='images/<%response.write web_var(web_config,5)%>/login_bg.gif'>
<%
  if login_username<>"" and login_mode<>"" then
%>
  <table border=0 cellspacing=0 cellpadding=0>
  <tr height=20><td></td></tr>
  <tr height=30><td align=center>你好，<b><font class=blue><%response.write login_username%></font></b></td></tr>
  <tr><td align=center>你现在已登陆 <font class=red><%response.write format_power(login_mode,1)%></font> 模式</td></tr>
  <tr height=30><td align=center><%
    dim mess_dim
    if login_message>0 then
      mess_dim="red"
      response.write "<bgsound src='images/mail/message.wav' border=0>"
    else
      login_message=0:mess_dim="gray"
    end if
    response.write "<a href='user_mail.asp'><img src='images/mail/new.gif' align=absmiddle border=0>&nbsp;我的短信箱（<font class="&mess_dim&">"&login_message&"&nbsp;新</font>）</a>"
%></td></tr>
  <tr><td align=center><a href='login.asp?action=logout'>退出登陆</a>&nbsp;┋&nbsp;<a href='user_main.asp'>用户中心</a></td></tr>
  </table>
<%
  else
%>
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
<% end if %>
</td></tr><tr><td height=5></td></tr></table>
<%
end sub
sub web_head(var1,var2,var3,var4,var5)
  dim ttt,ntt,user_ip,user_sys,rs,sql,wt
  user_ip=ip_sys(1,1):user_sys=ip_sys(3,0)
  wt=int(mid(web_setup,3,1))
  
  if web_login=1 and index_url<>"error" then
    if symbol_name(login_username)<>"yes" or symbol_ok(login_password)<>"yes" then
      login_mode="":login_username="":login_password=""
    else
      sql="select id,power,popedom,emoney from user_data where hidden=1 and username='"&login_username&"' and password='"&login_password&"'"
      set rs=conn.execute(sql)
      if rs.eof and rs.bof then
        response.Cookies(web_cookies)("login_username")=""
        response.Cookies(web_cookies)("login_password")=""
      else
        login_popedom=rs("popedom"):login_mode=rs("power"):login_emoney=rs("emoney")
      end if
      rs.close
    end if
    if wt=1 and login_mode="" then
      ttt=request.cookies(web_cookies)("guest_name")
      set rs=conn.execute("select l_id from user_login where l_username='"&ttt&"'")
      if rs.eof and rs.bof then
        ntt="游客"&Session.SessionID
        conn.execute("insert into user_login(l_username,l_type,l_where,l_tim_login,l_tim_end,l_ip,l_sys) values('"&ntt&"',1,'"&tit&"','"&now_time&"','"&now_time&"','"&user_ip&"','"&user_sys&"')")
        response.cookies(web_cookies)("guest_name")=ntt
      else
        conn.execute("update user_login set l_where='"&tit&"',l_tim_end='"&now_time&"' where l_id="&rs("l_id"))
      end if
      rs.close
    end if
    if (wt=1 or wt=2) and login_mode<>"" then
      login_message=0
      sql="select count(*) from user_mail where accept_u='"&login_username&"' and types=1 and isread=0"
      set rs=conn.execute(sql)
      if not(rs.eof and rs.bof) then login_message=int(rs(0))
      rs.close
      sql="select l_id from user_login where l_username='"&login_username&"'"
      set rs=conn.execute(sql)
      if rs.eof and rs.bof then
        conn.execute("insert into user_login(l_username,l_type,l_where,l_tim_login,l_tim_end,l_ip,l_sys) values('"&login_username&"',0,'"&tit&"','"&now_time&"','"&now_time&"','"&user_ip&"','"&user_sys&"')")
      else
        conn.execute("update user_login set l_where='"&tit&"',l_tim_end='"&now_time&"' where l_username='"&login_username&"'")
      end if
      rs.close
      response.cookies(web_cookies)("guest_name")=""
    end if
    
    set rs=conn.execute("select count(l_id) from user_login")
    if not(rs.eof and rs.bof) then online_num=rs(0)
    rs.close:set rs=nothing
    if wt<>1 and login_mode="" then online_num=online_num+1
    counter_s=counter_type("view","no")
    '---------------------------------Access 2000----------------------------------
    'sql="delete from user_login where DateDiff('n',l_tim_end,now())>"&int(web_var(web_num,13))
    'sql="delete from user_login where l_tim_end<"&DateAdd("n",-"&int(web_var(web_num,13))&",now())
    'conn.execute(sql)
    call cookies_yes()
  end if
  if page_power<>"" and format_page_power(login_mode)<>"yes" then call close_conn():call cookies_type("power")
  select case var1
  case 1
    if int(web_var_num(web_setup,1,1))<>0 and (login_username="" or login_mode="") then call close_conn():call cookies_type("login")
  case 2
    if login_username="" or login_mode="" then call close_conn():call cookies_type("login")
  end select
  select case var2
  case 1
    call close_conn()
  case 2
    exit sub
  end select
%>
<html>
<head>
<title><%response.write web_var(web_config,1) & " - " & tit%></title>
<meta name="Description"  content="Beyondest">
<meta name="keywords" content="最全的Beyond资料,最好的Beyond网站,asp,Beyondest,笼民,书记">
<meta name="author" content="笼民">
<meta http-equiv=Content-Type content=text/html; charset=gb2312>
<link rel=stylesheet href='include/beyondest.css' type=text/css>
<script language=javascript src='style/open_win.js'></script>
<script language=javascript src='style/mouse_on_title.js'></script>
</head>
<body topmargin=0 leftmargin=0 bgcolor=<%
response.write web_var(web_color,1)
ttt=web_var(web_config,7)
if ttt<>"" then
  response.write " background='images/"&web_var(web_config,5)&"/"&ttt&".gif'"
end if
%>><a name='top'></a><center>
<%if int(var3)=4 then exit sub%>
<%response.write lefter%>
<table border=0 width='100%' cellspacing=0 cellpadding=0>
<tr><td height=5 bgcolor=<% response.write web_var(web_color,3) %>></td></tr>
<tr><td height=2 bgcolor=<% response.write web_var(web_color,4) %>></td></tr>
<tr><td><%=kong%></td></tr>
<tr><td align=center>
  <table border=0 width='100%' cellspacing=0 cellpadding=0 class=tf>
  <tr align=center height=80>
  <td width='33%'><a href='<%=web_var(web_config,2)%>' target=_blank><img src='images/<%=web_var(web_config,5)%>/top_logo.gif' border=0></a></td>
  <td ><object classid="clsid:D27CDB6E-AE6D-11cf-96B8-444553540000" codebase="http://download.macromedia.com/pub/shockwave/cabs/flash/swflash.cab#version=6,0,29,0" width="468" height="60">
    <param name="movie" value="images/<%=web_var(web_config,5)%>/banner.swf">
    <param name="quality" value="high">
    <embed src="images/<%=web_var(web_config,1)%>/banner.swf" quality="high" pluginspage="http://www.macromedia.com/go/getflashplayer" type="application/x-shockwave-flash" width="468" height="60"></embed>
  </object></td>
<td align=right width='7%'>
    <table border=0 cellspacing=0 cellpadding=1 align=center>
    <tr height=20><td><a class=top href="javascript:window.external.AddFavorite('<%response.write web_var(web_config,2)%>','<%response.write web_var(web_config,1)%>')" style='target: ' _self?>加入收藏</a></td></tr>
    <tr height=20><td><a class=top href='help.asp'>帮助中心</a></td></tr>
    <tr height=20><td><a class=top href='mailto:plinq@live.com'>联系我们</a></td></tr>
    
    </table>
  </td>  

  </tr></table>
</td></tr>
<tr><td valign=middle>

  <table border=0 cellspacing=0 cellpadding=0 class=tf align=center background='images/<%=web_var(web_config,5)%>/menu_bg.gif'>
  <%
  response.write "<tr align=center height=35>"
   dim wdim
    wdim=split(web_menu,"|")
  for i=0 to ubound(wdim)
    response.write vbcrlf&"<td><a href='"&left(wdim(i),instr(wdim(i),":")-1)&".asp'>"&format_icon(left(wdim(i),instr(wdim(i),":")-1))&"</a>"&right(wdim(i),len(wdim(i))-instr(wdim(i),":"))&"</td>"
  next
  erase wdim
  response.write "</tr>"
%></table>
</td></tr>
<tr><td height=1 bgcolor=<%response.write web_var(web_color,3)%>></td></tr>
</table>
<table border=0 cellpadding=0 cellspacing=0 width='100%' align=center>
<tr><td align=left height=25 width=620>&nbsp;&nbsp;
今天是：<%response.write formatdatetime(now_time,1)&"&nbsp;"&WeekDayName(WeekDay(now_time))%>
&nbsp;&nbsp;您现在位于：&nbsp;<a href='./'><%response.write web_var(web_config,1)%></a>&nbsp;→&nbsp;
<%
  if tit_fir="" then
    response.write tit
  else
    response.write "<a href='"&index_url&".asp'>" & tit_fir & "</a>&nbsp;→&nbsp;" & tit
  end if
  response.write "</td><td width=1 background='images/"&web_var(web_config,5)&"/bg_dian2.gif'></td><td width=154 align=right bgcolor="&web_var(web_color,1)&">"

  if login_message>0 then
    response.write "<a href='user_mail.asp'><img src='images/mail/new.gif' align=absmiddle border=0>&nbsp;我的短信箱（<font class=red>"&login_message&"&nbsp;新</font>）</a>"
  else
%>
<marquee scrolldelay=120 scrollamount=4 onMouseOut="if (document.all!=null){this.start()}" onMouseOver="if (document.all!=null){this.stop()}"><script src='style/head_scroll.js'></script></marquee>
<%
  end if
  response.write "</td></tr></table>" & _
		 vbcrlf&"<table border=0 cellpadding=0 cellspacing=0 width='100%' align=center><tr><td background='images/"&web_var(web_config,5)&"/bg_dian.gif'></td></tr></table>"&righter&web_left(var3)
end sub
function web_left(lt)
  select case lt
  case 1
    web_left=vbcrlf&"<table border=0 width=775 cellspacing=0 cellpadding=0><tr valign=top align=center><td height=300 width=580>"
  case 2
    web_left=vbcrlf&"<table border=0 width=775 cellspacing=0 cellpadding=0><tr><td align=center>"
  case 3
    web_left=vbcrlf&"<table border=0 width=775 cellspacing=0 cellpadding=0><tr><td height=300 align=center>"
  case else
    web_left=vbcrlf&"<table border=0 width=775 cellspacing=0 cellpadding=0><tr valign=top align=center><td height=300 width=185 bgcolor="&web_var(web_color,6)&">"
  end select
  web_left=lefter&web_left
end function
sub web_center(ct)
  dim ttt,ttl,ttr:ttt="bg3"
  if ttt<>"" then ttr=" background='images/"&web_var(web_config,5)&"/"&ttt&"r.gif'":ttl=" background='images/"&web_var(web_config,5)&"/"&ttt&"l.gif'"
  select case ct
  case 1
    response.write vbcrlf&"</td>" & _
       vbcrlf&"<td width=1 bgcolor="&web_var(web_color,14)&"></td>" & _
       vbcrlf&"<td width=8 bgcolor="&web_var(web_color,5)&ttr&"></td>" & _
       vbcrlf&"<td width=1 bgcolor="&web_var(web_color,14)&"></td>" & _
       vbcrlf&"<td wdith=185 bgcolor="&web_var(web_color,6)&">"
  case else
    response.write vbcrlf&"</td>" & _
       vbcrlf&"<td width=1 bgcolor="&web_var(web_color,14)&"></td>" & _
       vbcrlf&"<td width=8 bgcolor="&web_var(web_color,5)&ttl&"></td>" & _
       vbcrlf&"<td width=1 bgcolor="&web_var(web_color,14)&"></td>" & _
       vbcrlf&"<td wdith=580>"
  end select
end sub
function web_right()
  web_right=vbcrlf&"</td></tr></table>"&righter
end function
sub web_end(wt)
  if isobject(rs) then set rs=nothing
  if wt=0 then call close_conn()
  response.write web_right()&lefter
%>
<table border=0 width='100%' cellspacing=0 cellpadding=0>
<tr><td height=1 bgcolor=<%response.write web_var(web_color,3)%>></td></tr>
<tr><td class=end align=center height=20 bgcolor=<% response.write web_var(web_color,2) %>>
<a class=menu href='help.asp?action=about'>关于我们</a>&nbsp;┋
<a class=menu href='gbook.asp'>网站留言</a>&nbsp;┋
<a href='links.asp' class=menu>友情链接</a>&nbsp;┋
<a class=menu href='online.asp'>与我在线</a>&nbsp;┋
版本：<a href='help.asp' target=_blank class=menu><% response.write web_edition %></a>&nbsp;┋
<a class=menu href='admin_login.asp' target=_blank>管理</a>&nbsp;┋
<a href='#top' class=menu>TOP</a>
</td></tr>
<tr><td height=1 bgcolor=<% response.write web_var(web_color,3) %>></td></tr>
<tr><td align=center height=20><% response.write counter_s%></td></tr>
<tr><td align=center><%response.write web_var(web_config,1)&"&nbsp;<font class=gray>"&web_var(web_stamp,int(mid(web_setup,3,1))+1)&"</font>"%>&nbsp;┋页面执行时间：<font class=red_3><% response.write FormatNumber((timer()-timer_start)*1000,3) %></font> 毫秒</td></tr>
<tr><td align=center height=20><% response.write web_var(web_error,4) %></td></tr>
<tr><td height=2 bgcolor=<% response.write web_var(web_color,4) %>></td></tr>
<tr><td height=5 bgcolor=<% response.write web_var(web_color,3)%>></td></tr>
</table><% response.write righter %>
</center></body></html>
<%
end sub
sub web_end2(wt)
  if isobject(rs) then set rs=nothing
  if wt=0 then call close_conn()
  response.write "</center></body></html>"
end sub
%>
<!-- #include file="INCLUDE/config_other.asp" -->
<!-- #include file="include/conn.asp" -->
<%
' ====================
' Beyondest.Com v3.6.1
' http://beyondest.com
' ====================

dim wt:wt=int(mid(web_setup,3,1))
tit="��������"

call web_head(0,0,0,0,0)
'------------------------------------left----------------------------------
call format_login()
response.write left_action("jt13",4)
'----------------------------------left end--------------------------------
call web_center(0)
'-----------------------------------center---------------------------------
response.write ukong
%>
<table border=0 width='98%'>
<tr><td align=center height=30><% call online_main() %></td></tr>
<tr><td align=center height=30><%response.write user_power_type(0)%></td></tr>
<tr><td align=center class=htd><%
select case wt
case 1
  response.write "<font class=red>��վ�����ѿ��� <font class=blue>"&web_var(web_stamp,wt+1)&"</font> ģʽ�����е�½�������վ���˱�����¼�����б�</font>"
case 2
  response.write "<font class=red>��վ�����ѿ��� <font class=blue>"&web_var(web_stamp,wt+1)&"</font> ģʽ����վ��ע���û����Ե�½������¼�û������б�</font>"
case else
  response.write "<font class=red>��վ�����ѿ��� <font class=blue>"&web_var(web_stamp,wt+1)&"</font> ģʽ����վ��ע���û����Ե�½������¼�����б�</font>"
end select

response.write "<br>�й� <a href='help.asp?action=web'>��վģʽ</a> ����ϸ˵��������� <a href='help.asp?action=web'>��վ����</a> �鿴�����Ϣ��"
%></td></tr>
<tr><td align=center height=5></td></tr>
</table>
<%
'---------------------------------center end-------------------------------
call web_end(0)

sub online_main()
  dim rssum,l_username
  if var_null(login_username)="" then
    response.write "<font class=blue>"
    if wt=1 then
      response.write request.cookies("beyondest_online")("guest_name")
    else
      response.write "�ο�"
    end if
    response.write "</font>�����ã�"&web_var(web_error,2)
  else
    response.write "<font class=blue>"&login_username&"</font>�����ã���ӭ��ע�Ტ��½��վ�������ڿ��Ե�����������Ŀ����ϸ���ݡ�"
  end if

  if wt=0 then
    response.write "<tr><td></td></tr><tr><td height=200>"
    exit sub
  end if

  set rs=server.createobject("adodb.recordset")
  
  if wt=1 or wt=2 then
    sql="select user_login.*,user_data.power from user_data inner join user_login on user_login.l_username=user_data.username where user_login.l_type=0 order by user_login.l_id"
    rs.open sql,conn,1,1
    if rs.eof and rs.bof then
      rssum=0
    else
      rssum=rs.recordcount
    end if
%>
</td></tr>
<tr><td height=5></td></tr>
<tr><td><% response.write img_small("jt1") %><font class=red_3><b>���߻�Ա</b></font>&nbsp;��<font class=red><% response.write rssum %></font>&nbsp;�ˣ�</td></tr>
<tr><td align=center height=150 valign=top>
  <table border=0 width='100%'>
  <tr><td width='20%'></td><td width='20%'></td><td width='20%'></td><td width='20%'></td><td width='20%'></td></tr>
<%
    do while not rs.eof
      response.write "<tr>"
      for i=1 to 5
        if rs.eof then exit for
        l_username=rs("l_username")
        response.write "<td>"&img_small("icon_"&rs("power"))&"<a href='user_view.asp?username="&server.urlencode(l_username)&"' title='Ŀǰλ�ã�"&rs("l_where")&"<br>����ʱ�䣺"&rs("l_tim_login")&"<br>�ʱ�䣺"&rs("l_tim_end")&"<br>��ʵIP��"&ip_types(rs("l_ip"),l_username,0)&"<br>"&view_sys(rs("l_sys"))&"' target=_blank>"&l_username&"</a></td>"
        rs.movenext
      next
      response.write "</tr>"
    loop
    rs.close
    response.write "</table>"
  end if
  
  if wt=1 then
    sql="select * from user_login where l_type=1 order by l_id"
    rs.open sql,conn,1,1
    if rs.eof and rs.bof then
      rssum=0
    else
      rssum=rs.recordcount
    end if
%>
</td></tr>
<tr><td><% response.write img_small("jt12") %><font class=red_3><b>�����ο�</b></font>&nbsp;��<font class=red><% response.write rssum %></font>&nbsp;�ˣ�</td></tr>
<tr><td align=center height=150 valign=top>
  <table border=0 width='98%'>
  <tr><td width='25%'></td><td width='25%'></td><td width='25%'></td><td width='25%'></td></tr>
<%
    do while not rs.eof
      response.write "<tr>"
      for i=1 to 4
        if rs.eof then exit for
        l_username=rs("l_username")
        response.write "<td>"&img_small("icon_other")&"<font title='Ŀǰλ�ã�"&rs("l_where")&"<br>����ʱ�䣺"&rs("l_tim_login")&"<br>�ʱ�䣺"&rs("l_tim_end")&"<br>"&ip_types(rs("l_ip"),l_username,0)&"<br>"&view_sys(rs("l_sys"))&"' target=_blank>"&l_username&"</font></td>"
        rs.movenext
      next
      response.write "</tr>"
    loop
    rs.close
    response.write "</table>"
  end if
  set rs=nothing
end sub
%>
<!-- #include file="INCLUDE/config_user.asp" -->
<!-- #include file="include/jk_ubb.asp" -->
<!-- #include file="include/conn.asp" -->
<%
' ====================
' Beyondest.Com v3.6.1
' http://beyondest.com
' ====================

dim integral,unit_num,emoney_1,chk,errs
tit="�������"

call web_head(2,0,0,0,0)
'------------------------------------left----------------------------------
call left_user()
'----------------------------------left end--------------------------------
call web_center(0)
'-----------------------------------center---------------------------------
response.write ukong
call emoney_top()

call emoney_main()

response.write ukong
'---------------------------------center end-------------------------------
call web_end(0)

sub emoney_main()
  dim emoneys,emoney_2,e_num,e_all,c_name,c_pass,c_emoney,c_id,userp
  unit_num=int(web_var(web_num,14)):errs="":emoney_2=0:c_id=0
  set rs=conn.execute("select integral from user_data where hidden=1 and username='"&login_username&"'")
  integral=rs("integral")
  rs.close:set rs=nothing
  emoney_1=integral\unit_num:userp=format_power(login_mode,2)
  if not(isnumeric(userp)) then userp=0
  userp=int(userp)
  chk=trim(request.querystring("chk"))
  if action<>"virement" and action<>"card" then action="converion"
  
  if (action="converion" or action="virement") and chk="yes" then
    e_num=trim(request.form("e_num")):e_all=trim(request.form("e_all"))
    emoneys=emoney_1
    if action="virement" then emoneys=login_emoney
    if e_all="yes" then
      emoney_2=emoneys
    else
      if not(isnumeric(e_num)) then
        errs="no"
      else
        if instr(1,e_num,".")>0 then
          errs="no"
        else
         if int(e_num)<1 or int(e_num)>int(emoneys) then
           errs="no"
         else
           emoney_2=e_num
         end if
        end if
      end if
    end if
    
    if action="converion" and int(emoney_2)>0 then
      conn.execute("update user_data set integral=integral-"&emoney_2*unit_num&",emoney=emoney+"&emoney_2&" where username='"&login_username&"'")
      integral=integral-emoney_2*unit_num:login_emoney=login_emoney+emoney_2:emoney_1=emoney_1-emoney_2
      response.write "<script language=javascript>alert(""���ѳɹ������� "&emoney_2&" "&m_unit&"��\n\n���Ļ��������ˣ�"&emoney_2*unit_num&" ��\n\nĿǰ�Ļ��ֻ�����Ϊ��ÿ "&unit_num&" �ֿɻ��� 1 "&m_unit&""");</script>"
    end if
    
    if action="virement" and int(emoney_2)>0 then
      dim username2:username2=trim(request.form("username2"))
      if symbol_name(username2)<>"yes" then
        errs="no"
      else
        set rs=conn.execute("select username from user_data where username='"&username2&"'")
        if rs.eof and rs.bof then errs="no"
        rs.close:set rs=nothing
      end if
      if errs="" then
        conn.execute("update user_data set emoney=emoney-"&emoney_2&" where username='"&login_username&"'")
        conn.execute("update user_data set emoney=emoney+"&emoney_2&" where username='"&username2&"'")
        login_emoney=login_emoney-emoney_2
        response.write "<script language=javascript>alert(""���ѳɹ��ĸ� "&username2&" ת���� "&emoney_2&" "&m_unit&"��\n\n����ӵ�е�"&tit&"Ҳ�����ˣ�"&emoney_2&" "&m_unit&""");</script>"
        sql="insert into user_mail(send_u,accept_u,topic,word,tim,types,isread) " & _
	    "values('"&login_username&"','"&username2&"','[ϵͳ]����ת����Ϣ��ʾ','"&login_username&" �ѳɹ��ĸ� �� ת���� "&emoney_2&" "&m_unit&"��','"&now_time&"',1,0)"
	conn.execute(sql)
      end if
    end if
  end if
  
  if action="card" and chk="yes" then
    c_name=code_form(trim(request.form("c_name")))
    c_pass=code_form(trim(request.form("c_pass")))
    if len(c_name)<1 or len(c_pass)<1 then errs="no"
    if errs="" then
      sql="select c_id,c_emoney from cards where c_name='"&c_name&"' and c_pass='"&c_pass&"' and c_hidden=0"
      set rs=conn.execute(sql)
      if rs.eof and rs.bof then
        errs="no"
      else
        c_id=rs("c_id"):c_emoney=rs("c_emoney")
      end if
      rs.close:set rs=nothing
    end if
    if errs="" then
      dim ok_msg:ok_msg=""
      conn.execute("update cards set c_hidden=1 where c_id="&c_id)
      sql="update user_data set emoney=emoney+"&c_emoney
      if int(userp)>3 then sql=sql&",power='"&format_power2(3,1)&"'":ok_msg="\n\n��Ҳͬʱ����Ϊ VIP ��Ա��"
      sql=sql&" where username='"&login_username&"'"
      conn.execute(sql)
      login_emoney=login_emoney+c_emoney
      response.write "<script language=javascript>alert(""���ѳɹ����û�Ա�������ţ�"&c_name&"��������ֵ�� "&c_emoney&" "&m_unit&"��"&ok_msg&""");</script>"
    end if
  end if
  
  select case action
  case "virement"
    call emoney_virement()
    call emoney_card()
    call emoney_converion()
  case "card"
    call emoney_card()
    call emoney_converion()
    call emoney_virement()
  case else
    call emoney_converion()
    call emoney_virement()
    call emoney_card()
  end select
  
  response.write ukong&table1
%>
<tr<%response.write table2%>><td>&nbsp;<%response.write img_small("fk00")%>&nbsp;<font class=end><b>���˵��</b></font></td></tr>
<tr<%response.write table3%>><td align=center>
  <table border=0 width='96%'>
  <tr><td height=25><font class=red>ע�⣺</font></td><td>������Ļ����<%response.write m_unit%>��ֵ���ܳ�����Ŀǰ���Ի�������ֵ��<font class=red><%response.write emoney_1&"</font>&nbsp;"&m_unit%>��</td></tr>
  <tr><td height=25></td><td>�������Ҫת�ʵ�<%response.write m_unit%>��ֵ���ܳ�����Ŀǰӵ�е����ֵ��<font class=red><%response.write login_emoney&"</font>&nbsp;"&m_unit%>��</td></tr>
  <tr><td height=25></td><td>����������е�<font class=blue>���ֻ���</font>��<font class=blue>����ת��</font>Ϊ<font class=red>���������</font>�����ڲ���ǰע��һ�¡�</td></tr>
  </table>
</td></tr>
</table><%
  response.write ukong
end sub

sub emoney_converion()
  response.write ukong&table1
%>
<tr<%response.write table2%>><td>&nbsp;<%response.write img_small("fk00")%>&nbsp;<font class=end><b>���ֻ���</b></font></td></tr>
<tr<%response.write table3%>><td align=center>
  <table border=0 width='96%'>
  <tr><td height=25>��Ŀǰӵ�е�<%response.write tit%>Ϊ��<font class=red><%response.write login_emoney&"</font>&nbsp;"&m_unit%></td></tr>
  </table>
</td></tr>
<tr<%response.write table3%>><td align=center>
  <table border=0 width='96%'>
  <tr><td height=25>Ŀǰ�Ļ��ֻ�����Ϊ��ÿ&nbsp;<font class=red_3><b><%response.write unit_num%></b></font>&nbsp;�ֿɻ���&nbsp;<font class=red><b>1</b></font>&nbsp;<%response.write m_unit%></td></tr>
  <tr><td height=25>��Ŀǰ����������Ϊ��<font class=red_3><%response.write integral%></font>&nbsp;��</td></tr>
  <tr><td height=25>��Ŀǰ���Ի��㣺<font class=red><%response.write emoney_1&"</font>&nbsp;"&m_unit%></td></tr>
  </table>
</td></tr>
<tr<%response.write table3%>><td align=center>
  <table border=0 width='96%'>
<% if action="converion" and chk="yes" and errs<>"" then %>
  <tr><td height=50><font class=red_2>����ʧ�ܣ�</font>������һ�������� <font class=red><%response.write emoney_1%></font> ����������
&nbsp;&nbsp;&nbsp;&nbsp;<%response.write go_back%></td></tr>
<% else %>
  <form name=emoney_frm_1 action='?action=converion&chk=yes' method=post>
  <tr><td height=50>��������Ҫ�����<%response.write m_unit%>��ֵ��&nbsp;
<input type=text name=e_num size=12 maxlength=10 value=''>&nbsp;&nbsp;&nbsp;
<input type=checkbox name=e_all value='yes'>&nbsp;ȫ������&nbsp;&nbsp;&nbsp;
<input type=submit value='���л���'></td></tr>
  </form>
<% end if %>
  </table>
</td></tr>
</table><%
end sub

sub emoney_virement()
  response.write ukong&table1
%>
<tr<%response.write table2%>><td>&nbsp;<%response.write img_small("fk00")%>&nbsp;<font class=end><b>����ת��</b></font></td></tr>
<tr<%response.write table3%>><td align=center>
  <table border=0 width='96%'>
  <tr><td height=25>��Ŀǰӵ�е�<%response.write tit%>Ϊ��<font class=red><%response.write login_emoney&"</font>&nbsp;"&m_unit%></td></tr>
  </table>
</td></tr>
<tr<%response.write table3%>><td align=center>
  <table border=0 width='96%'>
<% if action="virement" and chk="yes" and errs<>"" then %>
  <tr><td height=50><font class=red_2>ת��ʧ�ܣ�</font></td><td>������һ�������� <font class=red><%response.write emoney_1%></font> ��������&nbsp;��&nbsp;��Ҫת���ע���û������ڣ�&nbsp;&nbsp;<%response.write go_back%></td></tr>
<% else %>
  <form name=emoney_frm_2 action='?action=virement&chk=yes' method=post>
  <tr><td height=10></td></tr>
  <tr><td height=30>��������Ҫת�ʵ�ע���û���&nbsp;
<input type=text name=username2 size=15 maxlength=20 value=''>&nbsp;&nbsp;&nbsp;
<%response.write friend_select()%>
</td></tr>
  <tr><td height=30>��������Ҫת�ʵ�<%response.write m_unit%>��ֵ��&nbsp;
<input type=text name=e_num size=12 maxlength=10 value=''>&nbsp;&nbsp;&nbsp;
<input type=checkbox name=eall value='yes'>&nbsp;ȫ��ת��&nbsp;&nbsp;&nbsp;
<input type=submit value='����ת��'></td></tr>
  <tr><td height=10></td></tr>
  </form>
<% end if %>
  </table>
</td></tr>
</table><%
end sub

sub emoney_card()
  response.write ukong&table1
%>
<tr<%response.write table2%>><td>&nbsp;<%response.write img_small("fk00")%>&nbsp;<font class=end><b>��Ա����ֵ</b></font></td></tr>
<tr<%response.write table3%>><td align=center>
  <table border=0 width='96%'>
  <tr><td height=25>��Ŀǰӵ�е�<%response.write tit%>Ϊ��<font class=red><%response.write login_emoney&"</font>&nbsp;"&m_unit%></td></tr>
  </table>
</td></tr>
<tr<%response.write table3%>><td align=center>
  <table border=0 width='96%'>
<% if action="card" and chk="yes" and errs<>"" then %>
  <tr><td height=50><font class=red_2>��Ա����ֵʧ�ܣ�</font></td><td>������Ļ�Ա <font class=red>����</font> �� <font class=red>����</font> �д���&nbsp;&nbsp;<%response.write go_back%></td></tr>
<% else %>
  <form name=emoney_frm_3 action='?action=card&chk=yes' method=post>
  <tr><td height=50>
    <table border=0>
    <tr>
    <td>���ţ�&nbsp;<input type=text name=c_name size=15 maxlength=20></td>
    <td>&nbsp;&nbsp;���룺&nbsp;<input type=password name=c_pass size=15 maxlength=20></td>
    <td>&nbsp;&nbsp;<input type=submit value='��Ա����ֵ'></td>
    </tr>
    </table>
  </td><tr>
  </form>
<% end if %>
  </table>
</td></tr>
</table><%
end sub

sub emoney_top()
%>
<table border=0>
<tr align=center>
<td height=50><a href='?action=converion'><img src='IMAGES/SMALL/emoney_converion.gif' border=0></a></td>
<td width=50></td>
<td><a href='?action=virement'><img src='IMAGES/SMALL/emoney_virement.gif' border=0></a></td>
<td width=50></td>
<td><a href='?action=card'><img src='IMAGES/SMALL/emoney_card.gif' border=0></a></td>
</tr>
</table>
<%
end sub

function friend_select()
  dim sql,rs,ttt
  friend_select=vbcrlf&"<script language=javascript>" & _
		vbcrlf&"function Do_accept(addaccept) {" & _
		vbcrlf&"  if (addaccept!=0) { document.emoney_frm_2.username2.value=addaccept; }" & _
		vbcrlf&"  return;" & _
		vbcrlf&"}</script>" & _
		vbcrlf&"<select name=friend_select size=1 onchange=Do_accept(this.options[this.selectedIndex].value)>" & _
		vbcrlf&"<option value='0'>ѡ���ҵĺ���</option>"
  sql="select username2 from user_friend where username1='"&login_username&"' order by id"
  set rs=conn.execute(sql)
  do while not rs.eof
    ttt=rs(0)
    friend_select=friend_select&vbcrlf&"<option value='"&ttt&"'>"&ttt&"</option>"
    rs.movenext
  loop
  rs.close
  friend_select=friend_select&vbcrlf&"</select>"
end function
%>
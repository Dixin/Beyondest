<!-- #include file="config.asp" -->
<!-- #include file="skin.asp" -->
<!-- #include file="jk_md5.asp" -->
<%
' ====================
' Beyondest.Com v3.6.1
' http://beyondest.com
' ====================

dim err_head
err_head=img_small("jt0")
index_url="user_main"
tit_fir=format_menu(index_url)

sub nopass()
  dim pass_action
  pass_action=trim(request.form("pass_action"))
  select case pass_action
  case "question"
    if post_chk()<>"yes" then
      call close_conn()
      call cookies_type("post")
    end if
    response.write pass_question()
  case "chk"
    if post_chk()<>"yes" then
      call close_conn()
      call cookies_type("post")
    end if
    response.write pass_chk()
  case else
    response.write pass_type()
  end select
end sub

function pass_question()
  dim username
  username=trim(request.form("username"))
  if symbol_name(username)<>"yes" then
    pass_question="������� <font class=red>��½����</font> Ϊ�ջ򲻷�����ع���<br><br>"&go_back
    exit function
  end if
  pass_question="<table border=0 class=fr><form action='login.asp?action=nopass' method=post><input type=hidden name=pass_action value='chk'><tr height=40><td>��½���ƣ�</td><td><input type=text name=uname size=20 value='"&username&"' readonly class=black_bg></td></tr><tr height=25><td>����Կ�ף�</td><td><input type=password name=passwd size=20 maxlength=20></td></tr><tr height=25><td>�µ����룺</td><td><input type=password name=password size=20 maxlength=20></td></tr><tr height=25><td>�ظ����룺</td><td><input type=password name=password2 size=20 maxlength=20></td></tr><tr height=40><td></td><td><input type=submit value='�� һ ��'></td></tr><input type=hidden name=username value='"&username&"'></form></table>"
end function

function pass_chk()
  dim username,uname,passwd,password,password2
  username=trim(request.form("username"))
  uname=trim(request.form("uname"))
  passwd=trim(request.form("passwd"))
  password=trim(request.form("password"))
  password2=trim(request.form("password2"))
  if symbol_name(username)<>"yes" or username<>uname then
    pass_chk="������� <font class=red>��½����</font> Ϊ�ջ򲻷�����ع���<br><br>"&go_back
    exit function
  end if
  if symbol_name(passwd)<>"yes" then
    pass_chk="������� <font class=red>����Կ��</font> Ϊ�ջ򲻷�����ع���<br><br>"&go_back
    exit function
  end if
  if symbol_ok(password)<>"yes" then
    pass_chk="������� <font class=red>��½����</font> Ϊ�ջ򲻷�����ع���<br><br>"&go_back
    exit function
  else
    if password<>password2 then
      pass_chk="<font class=red>��½����</font> �� <font class=red>ȷ������</font> ��һ�£�<br><br>"&go_back
      exit function
    end if
  end if
  sql="select top 1 password from user_data where username='"&username&"' and passwd='"&jk_md5(passwd,"short")&"' and hidden=1"
  set rs=server.createobject("adodb.recordset")
  rs.open sql,conn,1,3
  if rs.eof and rs.bof then
    rs.close:set rs=nothing
    pass_chk="<font class=red>��½����</font> �� <font class=red>����Կ��</font> �д�����ѱ�������<br><br>"&go_back
    exit function
  end if
  rs("password")=jk_md5(password,"short")
  rs.update
  rs.close:set rs=nothing
  pass_chk="<font class=blue_1><b>"&username&"</b></font>��<font class=red>���ѳɹ��޸����������룡</font><br><br>�������ǣ�<font class=red_3>"&password2&"</font> ���ͼǣ�<br><br><a href='login.asp'>��������½ҳ��</a>"
end function

function pass_type()
  pass_type="<table border=0><form action='login.asp?action=nopass' method=post><input type=hidden name=pass_action value='question'><tr height=40><td>���ĵ�½���ƣ�</td><td><input type=text name=username size=20 maxlength=20></td></tr><tr height=40><td></td><td><input type=submit value='�� һ ��'></td></tr></form></table>"
end function

sub register_main()
  dim reg_action,left_i
  reg_action=trim(request.form("reg_action"))
  select case reg_action
  case "reg_main"
    left_i=2
  case "reg_chk"
    left_i=3
  case else
    left_i=1
  end select
%>
<table border=0 width='100%' cellspacing=0 cellpadding=0>
<tr valign=top align=center><td width='23%'>
<br><br><br><img name=reg_left src='images/<%response.write web_var(web_config,5)%>/reg_left_<%=left_i%>.gif' border=0>
</td><td width='77%'>
  <table border=0 width='90%' cellspacing=0 cellpadding=0>
  <tr><td align=center height=80><img src='images/<%response.write web_var(web_config,5)%>/reg_top.gif' border=0></td></tr>
  <tr><td align=center height=300><%
  select case reg_action
  case "reg_main"
    call reg_type()
  case "reg_chk"
    response.write reg_chk()
  case else
    call reg_policy()
  end select
%><br><br></td></tr>
  </table>
</td></tr></table>
<%
end sub

sub reg_policy()
%>
<table border=0 width=450 cellspacing=0 cellpadding=0>
<tr><td class=htd>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;��ӭ�����뱾վ��μӽ��������ۣ���վ�㽫������վ��ת�䡣<br><br>
Ϊά�����Ϲ������������ȶ��������Ծ������������<br><br>
��һ���������ñ�վΣ�����Ұ�ȫ��й¶�������ܣ������ַ�������Ἧ��ĺ͹���ĺϷ�Ȩ�棬�������ñ�վ���������ƺʹ���������Ϣ�� <br>
������һ��ɿ�����ܡ��ƻ��ܷ��ͷ��ɡ���������ʵʩ�ģ�<br>
����������ɿ���߸�������Ȩ���Ʒ���������ƶȵģ�<br>
����������ɿ�����ѹ��ҡ��ƻ�����ͳһ�ģ�<br>
�������ģ�ɿ�������ޡ��������ӣ��ƻ������Ž�ģ�<br>
�������壩�������������ʵ��ɢ��ҥ�ԣ������������ģ�<br>
��������������⽨���š����ࡢɫ�顢�Ĳ�����������ɱ���ֲ�����������ģ�<br>
�������ߣ���Ȼ�������˻���������ʵ�̰����˵ģ����߽����������⹥���ģ�<br>
�������ˣ��𺦹��һ��������ģ�<br>
�������ţ�����Υ���ܷ��ͷ�����������ģ�<br>
������ʮ��������ҵ�����Ϊ�ġ�<br>
�������������أ����Լ������ۺ���Ϊ����</td></tr>
<form name=form_reg action='login.asp?action=register' method=post>
<input type=hidden name=reg_action value='reg_main'>
<tr><td align=center height=50>
<input type=submit value="�����Ķ���ͬ����������">&nbsp;��&nbsp;<input type=button value="��ͬ��" onClick="document.location='index.asp'">
</td></tr>
</form>
</table>
<%
end sub

sub reg_type()
%><br>
  <table border=0 width=360 cellspacing=0 cellpadding=2>
  <tr><td width='35%'></td><td width='65%'></td></tr>
  <form name=reg_frm action='login.asp?action=register' method=post>
  <input type=hidden name=reg_action value='reg_chk'>
  <tr>
    <td align=center>�û����ƣ�</td>
    <td><input type=text name=username size=20 maxlength=20><%=redx%></td>
  </tr>
  <tr>
    <td align=center>��½���룺</td>
    <td><input type=password name=password size=20 maxlength=20><%=redx%></td>
  </tr>
  <tr>
    <td align=center>ȷ�����룺</td>
    <td><input type=password name=password2 size=20 maxlength=20><%=redx%></td>
  </tr>
  <tr>
    <td align=center>����Կ�ף�</td>
    <td><input type=text name=passwd size=20 maxlength=20><%=redx%></td>
  </tr>
  <tr>
    <td align=center>E - mail��</td>
    <td><input type=text name=email size=30 maxlength=50><%=redx%></td>
  </tr>
  <tr>
    <td align=center>�����Ա�</td>
    <td>&nbsp;<input type=radio name=sex value='boy' checked class=bg_1>&nbsp;�к�&nbsp;&nbsp;&nbsp;<input type=radio name=sex value='girl' class=bg_1>&nbsp;Ů��&nbsp;<%response.write redx%></td>
  </tr>
  <tr><td></td><td height=50><input type=submit value=' �� �� ע �� '></td></tr>
</form>
  <tr><td colspan=2 height=30><hr size=1 color=#c0c0c0 width='90%'></td></tr>
  <tr><td colspan=2>
<p style='line-height: 150%'><font class=red_2>�û�ע��ע�����</font><br>
&nbsp;&nbsp;&nbsp;1���û�����ע������ɹ�֮�󽫲��ܸ��ġ�<br>
&nbsp;&nbsp;&nbsp;2���û����ƿ����Ǵ�СдӢ����ĸ��a~z��A~Z�������������֣�0~9����
�����ַ���-�����»��ߡ�_���ͺ�����ɣ����ַ�����Ϊ�����ַ���-�����»��ߡ�_�������Ȳ��ܳ���20λ������joe_527<br>
&nbsp;&nbsp;&nbsp;3����½����ֻ���ɴ�СдӢ����ĸ��a~z��A~Z�������������֣�0~9����
�����ַ���-�����»��ߡ�_����ɡ�����dw7v9j<br>
&nbsp;&nbsp;&nbsp;4��������д��ע����Ϣ���ڿվ������ִ�Сд��</p>
  </td></tr>
  </table>
<%
end sub

function reg_chk()
  dim username,password,password2,passwd,email,red
  username=trim(request.form("username"))
  password=trim(request.form("password"))
  password2=trim(request.form("password2"))
  passwd=trim(request.form("passwd"))
  email=code_form(trim(request.form("email")))
  red=""
  if symbol_name(username)<>"yes" then
    red=red&err_head&"������� <font class=red>�û�����</font> Ϊ�ջ򲻷�����ع���<br>"
  else
    if health_name(username)<>"yes" then
      red=red&err_head&"������� <font class=red>�û�����</font> ����<font class=red>��ϵͳ�����ַ�</font>��<br>"
    end if
  end if
  if symbol_ok(password)<>"yes" then
    red=red&err_head&"������� <font class=red>��½����</font> Ϊ�ջ򲻷�����ع���<br>"
  else
    if password<>password2 then
      red=red&err_head&"������� <font class=red>��½����</font> �� <font class=red>ȷ������</font> ��һ�£�<br>"
    end if
  end if
  if symbol_name(passwd)<>"yes" then
    red=red&err_head&"������� <font class=red>����Կ��</font> Ϊ�ջ򲻷�����ع���<br>"
  end if
  if email_ok(email)<>"yes" or len(email)>50 then
    red=red&err_head&"������� <font class=red>E-mail</font> Ϊ�ջ򲻷����ʼ�����<br>"
  end if

  if red="" then
    sql="select * from user_data where username='" & username & "'"
    set rs=server.createobject("adodb.recordset")
    rs.open sql,conn,1,3
    if rs.eof and rs.bof then
      rs.addnew
      rs("username")=username
      rs("password")=jk_md5(password,"short")
      rs("passwd")=jk_md5(passwd,"short")
      rs("email")=email
      if trim(request.form("sex"))="girl" then
        rs("sex")=0
      else
        rs("sex")=1
      end if
      rs("face")="0"
      rs("tim")=now_time
      rs("power")="user"
      if web_var_num(web_setup,2,1)=0 then
        rs("hidden")=false
      else
        rs("hidden")=true
      end if
      rs("bbs_counter")=0
      rs("counter")=0
      rs("integral")=0
      rs("emoney")=0
      rs("login_num")=0
      rs("last_tim")=now_time
      rs("popedom")="00000000000000000000000000000000000000000000000000"
      rs.update
      rs.close:set rs=nothing
      
      conn.execute("update configs set new_username='"&username&"',num_reg=num_reg+1 where id=1")
      call reg_msg(username)
      
      if web_var_num(web_setup,2,1)=0 then
        reg_chk="<font class=red><b>"&username&"</b></font>�����ѳɹ�ע���Ϊ��վ�û���<br><br>�����ڵ�״̬Ϊ��<font class=red_3>δ���</font>����ȴ�����Ա����ˡ�лл��"
      else
        reg_chk="��ϲ��<font class=red><b>"&username&"</b></font>�����ѳɹ�ע���Ϊ��վ�û���<br><br><a href='login.asp'>���ڽ��е�½</a><br><br>�뾡���½���޸����ĸ������ϡ�"
      end if
      exit function
    else
      red=err_head&"������� <font class=red>�û����ƣ�<b>"&username&"</b>��</font> �Ѿ���ע�ᣡ<br>" & _
	       err_head&"������ѡ���������� <font class=red>�û�����</font> �Բ�����ע�ᣡ<br>"
      rs.close:set rs=nothing
      reg_chk=found_error(red,300):exit function
    end if
    rs.close:set rs=nothing
  else
    red=red&err_head&"������й� <a href='help.asp?action=register' class=red_3>�û�ע��ע������</a> ��������д��"
    reg_chk=found_error(red,280):exit function
  end if
end function

sub reg_msg(accept_u)
  dim msg_topic,msg_word
  msg_topic=web_var(web_config,1)&" ��ӭ���ĵ�����"
  msg_word=web_var(web_config,1)&"ȫ���û��͹�����Ա��ӭ���ĵ�����[br]" & _
	   "�����κ������뼰ʱ��ϵ���ǡ�[br]" & _
	   "�����κ�ʹ���ϵ�������鿴��վ������[br]" & _
	   "��л�����ʱ�վ��������һ��������������ϼ�԰��"
  sql="insert into user_mail(send_u,accept_u,topic,word,tim,types,isread) " & _
      "values('"&web_var(web_config,3)&"','"&accept_u&"','"&msg_topic&"','"&msg_word&"','"&now_time&"',1,0)"
  conn.execute(sql)
end sub

sub login_chk()
  dim username,password,red,id,power,hidden,face
  if symbol_name(login_username)="yes" then
    username=login_username
  else
    username=trim(request.form("username"))
  end if
  if symbol_ok(login_password)="yes" then
    password=login_password
  else
    password=trim(request.form("password"))
    password=jk_md5(password,"short")
  end if

  red=""
  if symbol_name(username)<>"yes" then
    red=red&err_head&"������� <font class=red_3>�û�����</font> Ϊ�ջ򲻷�����ع���<br>"
  end if
  if symbol_ok(password)<>"yes" then
    red=red&err_head&"������� <font class=red_3>��½����</font> Ϊ�ջ򲻷�����ع���<br>"
  end if

  if red="" then
    sql="select top 1 id,face,power,hidden from user_data where username='"&username&"' and password='"&password&"'"
    set rs=conn.execute(sql)
    if rs.eof and rs.bof then
      red=err_head&"������� <font class=red>�û�����</font> �� <font class=red>��½����</font>  �д���<br>" & _
	       err_head&"�����������Բ�������½��վ��"
      response.write found_error(red,260)
    else
      power=rs("power"):hidden=rs("hidden")
      if hidden=true then
        'response.cookies(web_cookies).path=web_path
        response.cookies(web_cookies)("login_username")=username
        response.cookies(web_cookies)("login_password")=password

        sql="update user_data set last_tim='"&now_time&"' where username='"&username&"'"
        conn.execute(sql)
        tit=request.cookies(web_cookies)("guest_name")
        if var_null(tit)<>"" then
          conn.execute("delete from user_login where l_username='"&tit&"'")
        end if
        
        if trim(request.form("memery_info"))="yes" then
	  response.cookies(web_cookies)("iscookies")="yes"
	  response.cookies(web_cookies).expires=date+365
        end if
	'----------------------------------------------------------------------------
	if trim(request.form("re_log"))="yes" then
	  call close_conn()
	  response.redirect request.servervariables("http_referer")
	  response.end
	end if
	'----------------------------------------------------------------------------
	response.write "<meta http-equiv='refresh' content='"&web_var(web_num,5)&"; url="&index_url&".asp'><br><br><br>��ã�<font class=red>"&username&"</font>&nbsp;�������� <font class=red>"&format_power(power,1)&"</font> ��½ģʽ<br><br>" & _
		       vbcrlf & "<a href='"&index_url&".asp'>����"&tit_fir&"</a>&nbsp;��&nbsp;<a href='login.asp?action=logout'>�˳����ε�½</a><br><br><br>"
      else
        response.write "<font class=red>�����û��ʺŻ�δ��ˣ�</font><br><br>�������Ա��ϵ��"
      end if
    end if
    rs.close
  else
    red=red&err_head&"������й� <a href='help.asp?action=register' class=red_3>�û�ע��ע������</a> ��������д��"
    response.write found_error(red,280)
  end if
end sub

sub login_main()
%>
<script language=javascript src='style/login.js'></script>
<table border=0>
<form name=login_frm method=post action='login.asp?action=login_chk' onsubmit="return login_true()">
<tr><td align=center height=30>�û����ƣ�&nbsp;<input type=text name=username size=15 maxlength=20 tabindex=1>&nbsp;&nbsp;</td></tr>
<tr><td align=center>��½���룺&nbsp;<input type=password name=password size=15 maxlength=20 tabindex=2>&nbsp;&nbsp;</td></tr>
<tr><td align=center height=30 align=center><input type=radio name=memery_info value='no' class=bg_1 checked>&nbsp;��½һ��&nbsp;
<input type=radio name=memery_info value='yes' class=bg_1>&nbsp;���õ�½</td></tr>
<tr><td align=center>
<input type=button value='ע ��' onClick="document.location='login.asp?action=register'">&nbsp;&nbsp;
<input type=button value='��������' onClick="document.location='login.asp?action=nopass'">&nbsp;&nbsp;
<input type=submit value='�� ½' tabindex=3>
</td></tr>
</table>
<%
end sub
%>
<!-- #include file="config.asp" -->
<!-- #include file="skin.asp" -->
<!-- #include file="jk_md5.asp" -->
<%
' ====================
' Beyondest.Com v3.6.1
' http://beyondest.com
' ====================

Dim err_head
err_head  = img_small("jt0")
index_url = "user_main"
tit_fir   = format_menu(index_url)

Sub nopass()
    Dim pass_action
    pass_action = Trim(Request.form("pass_action"))

    Select Case pass_action
        Case "question"

            If post_chk() <> "yes" Then
                Call close_conn()
                Call cookies_type("post")
            End If

            Response.Write pass_question()
        Case "chk"

            If post_chk() <> "yes" Then
                Call close_conn()
                Call cookies_type("post")
            End If

            Response.Write pass_chk()
        Case Else
            Response.Write pass_type()
    End Select

End Sub

Function pass_question()
    Dim username
    username          = Trim(Request.form("username"))

    If symbol_name(username) <> "yes" Then
        pass_question = "������� <font class=red>��½����</font> Ϊ�ջ򲻷�����ع���<br><br>" & go_back
        Exit Function
    End If

    pass_question = "<table border=0 class=fr><form action='login.asp?action=nopass' method=post><input type=hidden name=pass_action value='chk'><tr height=40><td>��½���ƣ�</td><td><input type=text name=uname size=20 value='" & username & "' readonly class=black_bg></td></tr><tr height=25><td>����Կ�ף�</td><td><input type=password name=passwd size=20 maxlength=20></td></tr><tr height=25><td>�µ����룺</td><td><input type=password name=password size=20 maxlength=20></td></tr><tr height=25><td>�ظ����룺</td><td><input type=password name=password2 size=20 maxlength=20></td></tr><tr height=40><td></td><td><input type=submit value='�� һ ��'></td></tr><input type=hidden name=username value='" & username & "'></form></table>"
End Function

Function pass_chk()
    Dim username,uname,passwd,password,password2
    username     = Trim(Request.form("username"))
    uname        = Trim(Request.form("uname"))
    passwd       = Trim(Request.form("passwd"))
    password     = Trim(Request.form("password"))
    password2    = Trim(Request.form("password2"))

    If symbol_name(username) <> "yes" Or username <> uname Then
        pass_chk = "������� <font class=red>��½����</font> Ϊ�ջ򲻷�����ع���<br><br>" & go_back
        Exit Function
    End If

    If symbol_name(passwd) <> "yes" Then
        pass_chk = "������� <font class=red>����Կ��</font> Ϊ�ջ򲻷�����ع���<br><br>" & go_back
        Exit Function
    End If

    If symbol_ok(password) <> "yes" Then
        pass_chk = "������� <font class=red>��½����</font> Ϊ�ջ򲻷�����ع���<br><br>" & go_back
        Exit Function
    Else

        If password <> password2 Then
            pass_chk = "<font class=red>��½����</font> �� <font class=red>ȷ������</font> ��һ�£�<br><br>" & go_back
            Exit Function
        End If

    End If

    sql    = "select top 1 password from user_data where username='" & username & "' and passwd='" & jk_md5(passwd,"short") & "' and hidden=1"
    Set rs = Server.CreateObject("adodb.recordset")
    rs.open sql,conn,1,3

    If rs.eof And rs.bof Then
        rs.Close:Set rs = Nothing
        pass_chk = "<font class=red>��½����</font> �� <font class=red>����Կ��</font> �д�����ѱ�������<br><br>" & go_back
        Exit Function
    End If

    rs("password") = jk_md5(password,"short")
    rs.update
    rs.Close:Set rs = Nothing
    pass_chk = "<font class=blue_1><b>" & username & "</b></font>��<font class=red>���ѳɹ��޸����������룡</font><br><br>�������ǣ�<font class=red_3>" & password2 & "</font> ���ͼǣ�<br><br><a href='login.asp'>��������½ҳ��</a>"
End Function

Function pass_type()
    pass_type = "<table border=0><form action='login.asp?action=nopass' method=post><input type=hidden name=pass_action value='question'><tr height=40><td>���ĵ�½���ƣ�</td><td><input type=text name=username size=20 maxlength=20></td></tr><tr height=40><td></td><td><input type=submit value='�� һ ��'></td></tr></form></table>"
End Function

Sub register_main()
    Dim reg_action,left_i
    reg_action = Trim(Request.form("reg_action"))

    Select Case reg_action
        Case "reg_main"
            left_i = 2
        Case "reg_chk"
            left_i = 3
        Case Else
            left_i = 1
    End Select %>
<table border=0 width='100%' cellspacing=0 cellpadding=0>
<tr valign=top align=center><td width='23%'>
<br><br><br><img name=reg_left src='images/<% Response.Write web_var(web_config,5) %>/reg_left_<% = left_i %>.gif' border=0>
</td><td width='77%'>
  <table border=0 width='90%' cellspacing=0 cellpadding=0>
  <tr><td align=center height=80><img src='images/<% Response.Write web_var(web_config,5) %>/reg_top.gif' border=0></td></tr>
  <tr><td align=center height=300><%

    Select Case reg_action
        Case "reg_main"
            Call reg_type()
        Case "reg_chk"
            Response.Write reg_chk()
        Case Else
            Call reg_policy()
    End Select %><br><br></td></tr>
  </table>
</td></tr></table>
<%
End Sub

Sub reg_policy() %>
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
End Sub

Sub reg_type() %><br>
  <table border=0 width=360 cellspacing=0 cellpadding=2>
  <tr><td width='35%'></td><td width='65%'></td></tr>
  <form name=reg_frm action='login.asp?action=register' method=post>
  <input type=hidden name=reg_action value='reg_chk'>
  <tr>
    <td align=center>�û����ƣ�</td>
    <td><input type=text name=username size=20 maxlength=20><% = redx %></td>
  </tr>
  <tr>
    <td align=center>��½���룺</td>
    <td><input type=password name=password size=20 maxlength=20><% = redx %></td>
  </tr>
  <tr>
    <td align=center>ȷ�����룺</td>
    <td><input type=password name=password2 size=20 maxlength=20><% = redx %></td>
  </tr>
  <tr>
    <td align=center>����Կ�ף�</td>
    <td><input type=text name=passwd size=20 maxlength=20><% = redx %></td>
  </tr>
  <tr>
    <td align=center>E - mail��</td>
    <td><input type=text name=email size=30 maxlength=50><% = redx %></td>
  </tr>
  <tr>
    <td align=center>�����Ա�</td>
    <td>&nbsp;<input type=radio name=sex value='boy' checked class=bg_1>&nbsp;�к�&nbsp;&nbsp;&nbsp;<input type=radio name=sex value='girl' class=bg_1>&nbsp;Ů��&nbsp;<% Response.Write redx %></td>
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
End Sub

Function reg_chk()
    Dim username,password,password2,passwd,email,red
    username  = Trim(Request.form("username"))
    password  = Trim(Request.form("password"))
    password2 = Trim(Request.form("password2"))
    passwd    = Trim(Request.form("passwd"))
    email     = code_form(Trim(Request.form("email")))
    red       = ""

    If symbol_name(username) <> "yes" Then
        red   = red & err_head & "������� <font class=red>�û�����</font> Ϊ�ջ򲻷�����ع���<br>"
    Else

        If health_name(username) <> "yes" Then
            red = red & err_head & "������� <font class=red>�û�����</font> ����<font class=red>��ϵͳ�����ַ�</font>��<br>"
        End If

    End If

    If symbol_ok(password) <> "yes" Then
        red = red & err_head & "������� <font class=red>��½����</font> Ϊ�ջ򲻷�����ع���<br>"
    Else

        If password <> password2 Then
            red = red & err_head & "������� <font class=red>��½����</font> �� <font class=red>ȷ������</font> ��һ�£�<br>"
        End If

    End If

    If symbol_name(passwd) <> "yes" Then
        red = red & err_head & "������� <font class=red>����Կ��</font> Ϊ�ջ򲻷�����ع���<br>"
    End If

    If email_ok(email) <> "yes" Or Len(email) > 50 Then
        red = red & err_head & "������� <font class=red>E-mail</font> Ϊ�ջ򲻷����ʼ�����<br>"
    End If

    If red = "" Then
        sql    = "select * from user_data where username='" & username & "'"
        Set rs = Server.CreateObject("adodb.recordset")
        rs.open sql,conn,1,3

        If rs.eof And rs.bof Then
            rs.addnew
            rs("username")     = username
            rs("password")     = jk_md5(password,"short")
            rs("passwd")     = jk_md5(passwd,"short")
            rs("email")     = email

            If Trim(Request.form("sex")) = "girl" Then
                rs("sex") = 0
            Else
                rs("sex") = 1
            End If

            rs("face")     = "0"
            rs("tim")     = now_time
            rs("power")     = "user"

            If web_var_num(web_setup,2,1) = 0 Then
                rs("hidden") = False
            Else
                rs("hidden") = True
            End If

            rs("bbs_counter")     = 0
            rs("counter")     = 0
            rs("integral")     = 0
            rs("emoney")     = 0
            rs("login_num")     = 0
            rs("last_tim")     = now_time
            rs("popedom")     = "00000000000000000000000000000000000000000000000000"
            rs.update
            rs.Close:Set rs = Nothing

            conn.execute("update configs set new_username='" & username & "',num_reg=num_reg+1 where id=1")
            Call reg_msg(username)

            If web_var_num(web_setup,2,1) = 0 Then
                reg_chk = "<font class=red><b>" & username & "</b></font>�����ѳɹ�ע���Ϊ��վ�û���<br><br>�����ڵ�״̬Ϊ��<font class=red_3>δ���</font>����ȴ�����Ա����ˡ�лл��"
            Else
                reg_chk = "��ϲ��<font class=red><b>" & username & "</b></font>�����ѳɹ�ע���Ϊ��վ�û���<br><br><a href='login.asp'>���ڽ��е�½</a><br><br>�뾡���½���޸����ĸ������ϡ�"
            End If

            Exit Function
        Else
            red = err_head & "������� <font class=red>�û����ƣ�<b>" & username & "</b>��</font> �Ѿ���ע�ᣡ<br>" & _
            err_head & "������ѡ���������� <font class=red>�û�����</font> �Բ�����ע�ᣡ<br>"
            rs.Close:Set rs = Nothing
            reg_chk = found_error(red,300):Exit Function
        End If

        rs.Close:Set rs = Nothing
    Else
        red     = red & err_head & "������й� <a href='help.asp?action=register' class=red_3>�û�ע��ע������</a> ��������д��"
        reg_chk = found_error(red,280):Exit Function
    End If

End Function

Sub reg_msg(accept_u)
    Dim msg_topic,msg_word
    msg_topic = web_var(web_config,1) & " ��ӭ���ĵ�����"
    msg_word  = web_var(web_config,1) & "ȫ���û��͹�����Ա��ӭ���ĵ�����[br]" & _
    "�����κ������뼰ʱ��ϵ���ǡ�[br]" & _
    "�����κ�ʹ���ϵ�������鿴��վ������[br]" & _
    "��л�����ʱ�վ��������һ��������������ϼ�԰��"
    sql = "insert into user_mail(send_u,accept_u,topic,word,tim,types,isread) " & _
    "values('" & web_var(web_config,3) & "','" & accept_u & "','" & msg_topic & "','" & msg_word & "','" & now_time & "',1,0)"
    conn.execute(sql)
End Sub

Sub login_chk()
    Dim username,password,red,id,power,hidden,face

    If symbol_name(login_username) = "yes" Then
        username = login_username
    Else
        username = Trim(Request.form("username"))
    End If

    If symbol_ok(login_password) = "yes" Then
        password = login_password
    Else
        password = Trim(Request.form("password"))
        password = jk_md5(password,"short")
    End If

    red     = ""

    If symbol_name(username) <> "yes" Then
        red = red & err_head & "������� <font class=red_3>�û�����</font> Ϊ�ջ򲻷�����ع���<br>"
    End If

    If symbol_ok(password) <> "yes" Then
        red = red & err_head & "������� <font class=red_3>��½����</font> Ϊ�ջ򲻷�����ع���<br>"
    End If

    If red = "" Then
        sql     = "select top 1 id,face,power,hidden from user_data where username='" & username & "' and password='" & password & "'"
        Set rs  = conn.execute(sql)

        If rs.eof And rs.bof Then
            red = err_head & "������� <font class=red>�û�����</font> �� <font class=red>��½����</font>  �д���<br>" & _
            err_head & "�����������Բ�������½��վ��"
            Response.Write found_error(red,260)
        Else
            power = rs("power"):hidden = rs("hidden")

            If hidden = True Then
                'response.cookies(web_cookies).path=web_path
                Response.cookies(web_cookies)("login_username") = username
                Response.cookies(web_cookies)("login_password") = password

                sql                                       = "update user_data set last_tim='" & now_time & "' where username='" & username & "'"
                conn.execute(sql)
                tit                                       = Request.cookies(web_cookies)("guest_name")

                If var_null(tit) <> "" Then
                    conn.execute("delete from user_login where l_username='" & tit & "'")
                End If

                If Trim(Request.form("memery_info")) = "yes" Then
                    Response.cookies(web_cookies)("iscookies") = "yes"
                    Response.cookies(web_cookies).expires     = Date + 365
                End If

                '----------------------------------------------------------------------------

                If Trim(Request.form("re_log")) = "yes" Then
                    Call close_conn()
                    Response.redirect Request.servervariables("http_referer")
                    Response.End
                End If

                '----------------------------------------------------------------------------
                Response.Write "<meta http-equiv='refresh' content='" & web_var(web_num,5) & "; url=" & index_url & ".asp'><br><br><br>��ã�<font class=red>" & username & "</font>&nbsp;�������� <font class=red>" & format_power(power,1) & "</font> ��½ģʽ<br><br>" & _
                vbcrlf & "<a href='" & index_url & ".asp'>����" & tit_fir & "</a>&nbsp;��&nbsp;<a href='login.asp?action=logout'>�˳����ε�½</a><br><br><br>"
            Else
                Response.Write "<font class=red>�����û��ʺŻ�δ��ˣ�</font><br><br>�������Ա��ϵ��"
            End If

        End If

        rs.Close
    Else
        red = red & err_head & "������й� <a href='help.asp?action=register' class=red_3>�û�ע��ע������</a> ��������д��"
        Response.Write found_error(red,280)
    End If

End Sub

Sub login_main() %>
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
End Sub %>
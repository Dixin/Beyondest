<!-- #include file="INCLUDE/config_user.asp" -->
<!-- #include file="include/jk_ubb.asp" -->
<!-- #include file="include/conn.asp" -->
<%
' ====================
' Beyondest.Com v3.6.1
' http://beyondest.com
' ====================

Dim integral,unit_num,emoney_1,chk,errs
tit = "�������"

Call web_head(2,0,0,0,0)
'------------------------------------left----------------------------------
Call left_user()
'----------------------------------left end--------------------------------
Call web_center(0)
'-----------------------------------center---------------------------------
Response.Write ukong
Call emoney_top()

Call emoney_main()

Response.Write ukong
'---------------------------------center end-------------------------------
Call web_end(0)

Sub emoney_main()
    Dim emoneys,emoney_2,e_num,e_all,c_name,c_pass,c_emoney,c_id,userp
    unit_num = Int(web_var(web_num,14)):errs = "":emoney_2 = 0:c_id = 0
    Set rs   = conn.execute("select integral from user_data where hidden=1 and username='" & login_username & "'")
    integral = rs("integral")
    rs.Close:Set rs = Nothing
    emoney_1 = integral\unit_num:userp = format_power(login_mode,2)
    If Not(IsNumeric(userp)) Then userp = 0
    userp    = Int(userp)
    chk      = Trim(Request.querystring("chk"))
    If action <> "virement" And action <> "card" Then action = "converion"

    If (action = "converion" Or action = "virement") And chk = "yes" Then
        e_num   = Trim(Request.form("e_num")):e_all = Trim(Request.form("e_all"))
        emoneys = emoney_1
        If action = "virement" Then emoneys = login_emoney

        If e_all = "yes" Then
            emoney_2 = emoneys
        Else

            If Not(IsNumeric(e_num)) Then
                errs = "no"
            Else

                If InStr(1,e_num,".") > 0 Then
                    errs = "no"
                Else

                    If Int(e_num) < 1 Or Int(e_num) > Int(emoneys) Then
                        errs     = "no"
                    Else
                        emoney_2 = e_num
                    End If

                End If

            End If

        End If

        If action = "converion" And Int(emoney_2) > 0 Then
            conn.execute("update user_data set integral=integral-" & emoney_2*unit_num & ",emoney=emoney+" & emoney_2 & " where username='" & login_username & "'")
            integral = integral - emoney_2*unit_num:login_emoney = login_emoney + emoney_2:emoney_1 = emoney_1 - emoney_2
            Response.Write "<script language=javascript>alert(""���ѳɹ������� " & emoney_2 & " " & m_unit & "��\n\n���Ļ��������ˣ�" & emoney_2*unit_num & " ��\n\nĿǰ�Ļ��ֻ�����Ϊ��ÿ " & unit_num & " �ֿɻ��� 1 " & m_unit & """);</script>"
        End If

        If action = "virement" And Int(emoney_2) > 0 Then
            Dim username2:username2 = Trim(Request.form("username2"))

            If symbol_name(username2) <> "yes" Then
                errs   = "no"
            Else
                Set rs = conn.execute("select username from user_data where username='" & username2 & "'")
                If rs.eof And rs.bof Then errs = "no"
                rs.Close:Set rs = Nothing
            End If

            If errs = "" Then
                conn.execute("update user_data set emoney=emoney-" & emoney_2 & " where username='" & login_username & "'")
                conn.execute("update user_data set emoney=emoney+" & emoney_2 & " where username='" & username2 & "'")
                login_emoney = login_emoney - emoney_2
                Response.Write "<script language=javascript>alert(""���ѳɹ��ĸ� " & username2 & " ת���� " & emoney_2 & " " & m_unit & "��\n\n����ӵ�е�" & tit & "Ҳ�����ˣ�" & emoney_2 & " " & m_unit & """);</script>"
                sql          = "insert into user_mail(send_u,accept_u,topic,word,tim,types,isread) " & _
                "values('" & login_username & "','" & username2 & "','[ϵͳ]����ת����Ϣ��ʾ','" & login_username & " �ѳɹ��ĸ� �� ת���� " & emoney_2 & " " & m_unit & "��','" & now_time & "',1,0)"
                conn.execute(sql)
            End If

        End If

    End If

    If action = "card" And chk = "yes" Then
        c_name = code_form(Trim(Request.form("c_name")))
        c_pass = code_form(Trim(Request.form("c_pass")))
        If Len(c_name) < 1 Or Len(c_pass) < 1 Then errs = "no"

        If errs = "" Then
            sql      = "select c_id,c_emoney from cards where c_name='" & c_name & "' and c_pass='" & c_pass & "' and c_hidden=0"
            Set rs   = conn.execute(sql)

            If rs.eof And rs.bof Then
                errs = "no"
            Else
                c_id = rs("c_id"):c_emoney = rs("c_emoney")
            End If

            rs.Close:Set rs = Nothing
        End If

        If errs = "" Then
            Dim ok_msg:ok_msg = ""
            conn.execute("update cards set c_hidden=1 where c_id=" & c_id)
            sql          = "update user_data set emoney=emoney+" & c_emoney
            If Int(userp) > 3 Then sql = sql & ",power='" & format_power2(3,1) & "'":ok_msg = "\n\n��Ҳͬʱ����Ϊ VIP ��Ա��"
            sql          = sql & " where username='" & login_username & "'"
            conn.execute(sql)
            login_emoney = login_emoney + c_emoney
            Response.Write "<script language=javascript>alert(""���ѳɹ����û�Ա�������ţ�" & c_name & "��������ֵ�� " & c_emoney & " " & m_unit & "��" & ok_msg & """);</script>"
        End If

    End If

    Select Case action
        Case "virement"
            Call emoney_virement()
            Call emoney_card()
            Call emoney_converion()
        Case "card"
            Call emoney_card()
            Call emoney_converion()
            Call emoney_virement()
        Case Else
            Call emoney_converion()
            Call emoney_virement()
            Call emoney_card()
    End Select

    Response.Write ukong & table1 %>
<tr<% Response.Write table2 %>><td>&nbsp;<% Response.Write img_small("fk00") %>&nbsp;<font class=end><b>���˵��</b></font></td></tr>
<tr<% Response.Write table3 %>><td align=center>
  <table border=0 width='96%'>
  <tr><td height=25><font class=red>ע�⣺</font></td><td>������Ļ����<% Response.Write m_unit %>��ֵ���ܳ�����Ŀǰ���Ի�������ֵ��<font class=red><% Response.Write emoney_1 & "</font>&nbsp;" & m_unit %>��</td></tr>
  <tr><td height=25></td><td>�������Ҫת�ʵ�<% Response.Write m_unit %>��ֵ���ܳ�����Ŀǰӵ�е����ֵ��<font class=red><% Response.Write login_emoney & "</font>&nbsp;" & m_unit %>��</td></tr>
  <tr><td height=25></td><td>����������е�<font class=blue>���ֻ���</font>��<font class=blue>����ת��</font>Ϊ<font class=red>���������</font>�����ڲ���ǰע��һ�¡�</td></tr>
  </table>
</td></tr>
</table><%
    Response.Write ukong
End Sub

Sub emoney_converion()
    Response.Write ukong & table1 %>
<tr<% Response.Write table2 %>><td>&nbsp;<% Response.Write img_small("fk00") %>&nbsp;<font class=end><b>���ֻ���</b></font></td></tr>
<tr<% Response.Write table3 %>><td align=center>
  <table border=0 width='96%'>
  <tr><td height=25>��Ŀǰӵ�е�<% Response.Write tit %>Ϊ��<font class=red><% Response.Write login_emoney & "</font>&nbsp;" & m_unit %></td></tr>
  </table>
</td></tr>
<tr<% Response.Write table3 %>><td align=center>
  <table border=0 width='96%'>
  <tr><td height=25>Ŀǰ�Ļ��ֻ�����Ϊ��ÿ&nbsp;<font class=red_3><b><% Response.Write unit_num %></b></font>&nbsp;�ֿɻ���&nbsp;<font class=red><b>1</b></font>&nbsp;<% Response.Write m_unit %></td></tr>
  <tr><td height=25>��Ŀǰ����������Ϊ��<font class=red_3><% Response.Write integral %></font>&nbsp;��</td></tr>
  <tr><td height=25>��Ŀǰ���Ի��㣺<font class=red><% Response.Write emoney_1 & "</font>&nbsp;" & m_unit %></td></tr>
  </table>
</td></tr>
<tr<% Response.Write table3 %>><td align=center>
  <table border=0 width='96%'>
<% If action = "converion" And chk = "yes" And errs <> "" Then %>
  <tr><td height=50><font class=red_2>����ʧ�ܣ�</font>������һ�������� <font class=red><% Response.Write emoney_1 %></font> ����������
&nbsp;&nbsp;&nbsp;&nbsp;<% Response.Write go_back %></td></tr>
<% Else %>
  <form name=emoney_frm_1 action='?action=converion&chk=yes' method=post>
  <tr><td height=50>��������Ҫ�����<% Response.Write m_unit %>��ֵ��&nbsp;
<input type=text name=e_num size=12 maxlength=10 value=''>&nbsp;&nbsp;&nbsp;
<input type=checkbox name=e_all value='yes'>&nbsp;ȫ������&nbsp;&nbsp;&nbsp;
<input type=submit value='���л���'></td></tr>
  </form>
<% End If %>
  </table>
</td></tr>
</table><%

End Sub

Sub emoney_virement()
Response.Write ukong & table1 %>
<tr<% Response.Write table2 %>><td>&nbsp;<% Response.Write img_small("fk00") %>&nbsp;<font class=end><b>����ת��</b></font></td></tr>
<tr<% Response.Write table3 %>><td align=center>
  <table border=0 width='96%'>
  <tr><td height=25>��Ŀǰӵ�е�<% Response.Write tit %>Ϊ��<font class=red><% Response.Write login_emoney & "</font>&nbsp;" & m_unit %></td></tr>
  </table>
</td></tr>
<tr<% Response.Write table3 %>><td align=center>
  <table border=0 width='96%'>
<% If action = "virement" And chk = "yes" And errs <> "" Then %>
  <tr><td height=50><font class=red_2>ת��ʧ�ܣ�</font></td><td>������һ�������� <font class=red><% Response.Write emoney_1 %></font> ��������&nbsp;��&nbsp;��Ҫת���ע���û������ڣ�&nbsp;&nbsp;<% Response.Write go_back %></td></tr>
<% Else %>
  <form name=emoney_frm_2 action='?action=virement&chk=yes' method=post>
  <tr><td height=10></td></tr>
  <tr><td height=30>��������Ҫת�ʵ�ע���û���&nbsp;
<input type=text name=username2 size=15 maxlength=20 value=''>&nbsp;&nbsp;&nbsp;
<% Response.Write friend_select() %>
</td></tr>
  <tr><td height=30>��������Ҫת�ʵ�<% Response.Write m_unit %>��ֵ��&nbsp;
<input type=text name=e_num size=12 maxlength=10 value=''>&nbsp;&nbsp;&nbsp;
<input type=checkbox name=eall value='yes'>&nbsp;ȫ��ת��&nbsp;&nbsp;&nbsp;
<input type=submit value='����ת��'></td></tr>
  <tr><td height=10></td></tr>
  </form>
<% End If %>
  </table>
</td></tr>
</table><%

End Sub

Sub emoney_card()
Response.Write ukong & table1 %>
<tr<% Response.Write table2 %>><td>&nbsp;<% Response.Write img_small("fk00") %>&nbsp;<font class=end><b>��Ա����ֵ</b></font></td></tr>
<tr<% Response.Write table3 %>><td align=center>
  <table border=0 width='96%'>
  <tr><td height=25>��Ŀǰӵ�е�<% Response.Write tit %>Ϊ��<font class=red><% Response.Write login_emoney & "</font>&nbsp;" & m_unit %></td></tr>
  </table>
</td></tr>
<tr<% Response.Write table3 %>><td align=center>
  <table border=0 width='96%'>
<% If action = "card" And chk = "yes" And errs <> "" Then %>
  <tr><td height=50><font class=red_2>��Ա����ֵʧ�ܣ�</font></td><td>������Ļ�Ա <font class=red>����</font> �� <font class=red>����</font> �д���&nbsp;&nbsp;<% Response.Write go_back %></td></tr>
<% Else %>
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
<% End If %>
  </table>
</td></tr>
</table><%

End Sub

Sub emoney_top() %>
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
End Sub

Function friend_select()
Dim sql,rs,ttt
friend_select = vbcrlf & "<script language=javascript>" & _
vbcrlf & "function Do_accept(addaccept) {" & _
vbcrlf & "  if (addaccept!=0) { document.emoney_frm_2.username2.value=addaccept; }" & _
vbcrlf & "  return;" & _
vbcrlf & "}</script>" & _
vbcrlf & "<select name=friend_select size=1 onchange=Do_accept(this.options[this.selectedIndex].value)>" & _
vbcrlf & "<option value='0'>ѡ���ҵĺ���</option>"
sql           = "select username2 from user_friend where username1='" & login_username & "' order by id"
Set rs        = conn.execute(sql)

Do While Not rs.eof
ttt           = rs(0)
friend_select = friend_select & vbcrlf & "<option value='" & ttt & "'>" & ttt & "</option>"
rs.movenext
Loop

rs.Close
friend_select = friend_select & vbcrlf & "</select>"
End Function %>
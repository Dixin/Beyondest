<!-- #include file="include/config_user.asp" -->
<%
' ====================
' Beyondest.Com v3.6.1
' http://beyondest.com
' ====================

Dim id:id = Trim(Request.querystring("id"))
If Not(IsNumeric(id)) And action <> "write" Then Call cookies_type("mail_id") %>
<!-- #include file="include/jk_ubb.asp" -->
<!-- #include file="include/conn.asp" -->
<%
Dim send_u
Dim accept_u
Dim topic
Dim word
Dim types
Dim isread
Dim red_3
tit = "վ�ڶ���"

Call web_head(2,0,0,0,0)

If action <> "view" And Int(popedom_format(login_popedom,41)) Then Call close_conn():Call cookies_type("locked")
'------------------------------------left----------------------------------
Call left_user()
'----------------------------------left end--------------------------------
Call web_center(0)
'-----------------------------------center---------------------------------
Response.Write ukong
Call user_mail_menu(0)
Response.Write ukong & table1

Select Case action
    Case "reply"
        Call mail_reply()
    Case "fw"
        Call mail_fw()
    Case "edit"
        Call mail_edit()
    Case "view"
        Response.Write mail_view()
    Case "del"
        Response.Write mail_del()
    Case Else
        Call mail_write()
End Select

Response.Write vbcrlf & "</table>"
'---------------------------------center end-------------------------------
Call web_end(0)

Function mail_del()
    mail_del = vbcrlf & "<tr" & table2 & "><td align=center><font class=end><b>ɾ������</b></font></td></tr>"
    Dim rs
    Dim sql
    Dim html_temp
    html_temp     = ""
    sql           = "select id from user_mail where (send_u='" & login_username & "' or accept_u='" & login_username & "') and id=" & id
    Set rs        = conn.execute(sql)

    If rs.eof And rs.bof Then
        html_temp = "<font class=red_2>����Ҫɾ���Ķ���ID�����ڻ����</font><br><br>" & go_back
    End If

    rs.Close:Set rs = Nothing

    If html_temp = "" Then
        sql       = "update user_mail set types=4 where id=" & id
        conn.execute(sql)
        html_temp = "<font class=red>����ɾ���ɹ���ɾ���Ķ��Ž��������Ļ���վ�ڡ�</font><br><br><a href='user_mail.asp?action=recycle'>�������</a>"
    End If

    mail_del      = mail_del & "<tr" & table3 & "><td height=150 align=center>" & html_temp & "</td></tr>"
End Function

Sub mail_write()
    Response.Write vbcrlf & "<tr" & table2 & " height=25><td colspan=2 align=center background=images/" & web_var(web_config,5) & "/bar_3_bg.gif><font class=end><b>׫д����</b></font></td></tr>"

    If Trim(Request.form("write_ok")) = "ok" Then
        Response.Write vbcrlf & "<tr" & table3 & "><td colspan=2 align=center height=150>"

        If post_chk() = "no" Then
            Response.Write web_var(web_error,1)
        Else
            red_3         = ""
            accept_u      = Trim(Request.form("accept_u"))
            topic         = Trim(Request.form("topic"))
            word          = Request.form("word")

            If symbol_name(accept_u) <> "yes" Then
                red_3     = red_3 & "<br><li><font class=red_3>�� �� ��</font> Ϊ�ջ򲻷�����ع���"
            Else
                sql       = "select username from user_data where username='" & accept_u & "'"
                Set rs    = conn.execute(sql)

                If rs.eof And rs.bof Then
                    red_3 = red_3 & "<br><li>����д�� <font class=red_3>�� �� ��</font> ���񲻴��ڣ�"
                End If

                rs.Close:Set rs = Nothing
            End If

            If var_null(topic) = "" Or Len(topic) > 20 Then
                red_3 = red_3 & "<br><li><font class=red_3>��������</font> ����Ϊ���ҳ��Ȳ��ܴ���20��"
            End If

            If var_null(word) = "" Or Len(word) > 250 Then
                red_3 = red_3 & "<br><li><font class=red_3>��������</font> ����Ϊ���ҳ��Ȳ��ܴ���250��"
            End If

            If red_3 = "" Then
                Set rs = Server.CreateObject("adodb.recordset")
                sql    = "select * from user_mail"
                rs.open sql,conn,1,3
                rs.addnew
                rs("send_u")     = login_username
                rs("accept_u")     = accept_u
                rs("topic")     = topic
                rs("word")     = word
                rs("tim")     = now_time

                If Trim(Request.form("send_later")) = "yes" Then
                    rs("types") = 2
                Else
                    rs("types") = 1
                End If

                rs("isread")     = False
                rs.update
                rs.Close

                If Trim(Request.form("send_later")) = "yes" Then
                    Response.Write "<font class=red>���ѳɹ��ı�����һ�����ţ�</font><br><br><a href='user_mail.asp?action=outbox'>�������</a>"
                Else
                    Response.Write "<font class=red>���ѳɹ��ĸ� <font class=blue><b>" & accept_u & "</b></font> ������һ�����ţ�</font><br><br><a href='user_mail.asp'>�������</a>"
                End If

            Else
                Response.Write found_error(red_3,"250")
            End If

        End If

        Response.Write vbcrlf & "</td></tr>"
    Else
        Response.Write vbcrlf & "<form name=mail_frm action='user_message.asp?action=write' method=post onsubmit=""javascript:frm_submitonce(this);""><input type=hidden name=write_ok value='ok'><input type=hidden name=send_later value=''>" & _
        vbcrlf & "<tr height=30" & table3 & "><td width='15%' align=center bgcolor=" & web_var(web_color,6) & ">�� �� �ˣ�</td><td width='85%'>&nbsp;<input type=text name=accept_u value='" & Trim(Request.querystring("accept_uaername")) & "' size=30 maxlength=20>" & redx & "&nbsp;��&nbsp;" & friend_select() & "</td></tr>" & _
        vbcrlf & "<tr height=30" & table3 & "><td align=center bgcolor=" & web_var(web_color,6) & ">�������⣺</td><td>&nbsp;<input type=text name=topic size=60 maxlength=20></td></tr>" & _
        vbcrlf & "<tr height=100" & table3 & "><td align=center class=htd bgcolor=" & web_var(web_color,6) & ">�������ݣ�<br>" & web_var(web_error,3) & "</td><td>&nbsp;<textarea cols=64 rows=6 name=word title='�����������250���ַ�<br>�� Ctrl+Enter ��ֱ�ӷ���' onkeydown=""javascript:frm_quicksubmit();""></textarea></td></tr>" & _
        vbcrlf & "<tr" & table3 & "><td colspan=2 height=40 align=center><input type=Submit name=wsubmit value='���Ͷ���'>&nbsp;��&nbsp;<input type=submit name=send value='�������' onclick=""javascript:mail_send_later();"">&nbsp;��&nbsp;<input type=reset value='�����д'></td></tr></form>"
    End If

End Sub

Sub mail_reply()
    Response.Write vbcrlf & "<tr" & table2 & "><td colspan=2 align=center><font class=end><b>�ظ�����</b></font></td></tr>"

    If Trim(Request.form("reply_ok")) = "ok" Then
        Response.Write vbcrlf & "<tr" & table3 & "><td colspan=2 align=center height=150>"

        If post_chk() = "no" Then
            Response.Write web_var(web_error,1)
        Else
            red_3         = ""
            accept_u      = Trim(Request.form("accept_u"))
            topic         = Trim(Request.form("topic"))
            word          = Request.form("word")

            If symbol_name(accept_u) <> "yes" Then
                red_3     = red_3 & "<br><li><font class=red_3>�� �� ��</font> Ϊ�ջ򲻷�����ع���"
            Else
                sql       = "select username from user_data where username='" & accept_u & "'"
                Set rs    = conn.execute(sql)

                If rs.eof And rs.bof Then
                    red_3 = red_3 & "<br><li>����д�� <font class=red_3>�� �� ��</font> ���񲻴��ڣ�"
                End If

                rs.Close
            End If

            If var_null(topic) = "" Or Len(topic) > 20 Then
                red_3 = red_3 & "<br><li><font class=red_3>��������</font> ����Ϊ���ҳ��Ȳ��ܴ���20��"
            End If

            If var_null(word) = "" Or Len(word) > 250 Then
                red_3 = red_3 & "<br><li><font class=red_3>��������</font> ����Ϊ���ҳ��Ȳ��ܴ���250��"
            End If

            If red_3 = "" Then
                Set rs = Server.CreateObject("adodb.recordset")
                sql    = "select * from user_mail"
                rs.open sql,conn,1,3
                rs.addnew
                rs("send_u")     = login_username
                rs("accept_u")     = accept_u
                rs("topic")     = topic
                rs("word")     = word
                rs("tim")     = now_time

                If Trim(Request.form("send_later")) = "yes" Then
                    rs("types") = 2
                Else
                    rs("types") = 1
                End If

                rs("isread")     = False
                rs.update
                rs.Close

                If Trim(Request.form("send_later")) = "yes" Then
                    Response.Write "<font class=red>���ѳɹ��ı�����һ�����ŵ����ݣ�</font><br><br><a href='user_mail.asp?action=outbox'>�������</a>"
                Else
                    Response.Write "<font class=red>���ѳɹ��ĸ� <font class=blue_1><b>" & accept_u & "</b></font> �ظ���һ�����ţ�</font><br><br><a href='user_mail.asp'>�������</a>"
                End If

            Else
                Response.Write found_error(red_3,"250")
            End If

        End If

        Response.Write vbcrlf & "</td></tr>"
    Else
        sql    = "select send_u,topic from user_mail where (send_u='" & login_username & "' or accept_u='" & login_username & "') and id=" & id
        Set rs = conn.execute(sql)

        If rs.eof And rs.bof Then
            rs.Close
            red_3 = "<br><li>�����ظ��� <font class=red_3>����ID</font> �����ڻ��д���"
            red_3 = found_error(red_3,"240")
            Response.Write vbcrlf & "<tr" & table3 & "><td align=center height=150 colspan=2>" & red_3 & "</td></tr>"

            Exit Sub
            Else
                Response.Write vbcrlf & "<form name=mail_frm action='user_message.asp?action=reply&id=" & id & "' method=post onsubmit=""javascript:frm_submitonce(this);""><input type=hidden name=reply_ok value='ok'><input type=hidden name=send_later value=''>" & _
                vbcrlf & "<tr height=30" & table3 & "><td width='15%' align=center>�� �� �ˣ�</td><td width='85%'>&nbsp;<input type=text name=accept_u value='" & rs("send_u") & "' size=30 maxlength=20>" & redx & "&nbsp;��&nbsp;" & friend_select() & "</td></tr>" & _
                vbcrlf & "<tr height=30" & table3 & "><td align=center>�������⣺</td><td>&nbsp;<input type=text name=topic value='RE:" & rs("topic") & "' size=60 maxlength=20></td></tr>" & _
                vbcrlf & "<tr height=100" & table3 & "><td align=center class=htd>�������ݣ�<br>" & web_var(web_error,3) & "</td><td>&nbsp;<textarea cols=64 rows=6 name=word title='�����������250���ַ�<br>�� Ctrl+Enter ��ֱ�ӷ���' onkeydown=""javascript:frm_quicksubmit();""></textarea></td></tr>" & _
                vbcrlf & "<tr" & table3 & "><td colspan=2 height=40 align=center><input type=Submit name=wsubmit value='���Ͷ���'>&nbsp;��&nbsp;<input type=Submit name=send value='�������' onclick=""javascript:mail_send_later();"">&nbsp;��&nbsp;<input type=reset value='�����д'></td></tr></form>"
            End If

            rs.Close
        End If

    End Sub

    Sub mail_fw()
        Response.Write vbcrlf & "<tr" & table2 & "><td colspan=2 align=center><font class=end><b>ת������</b></font></td></tr>"

        If Trim(Request.form("fw_ok")) = "ok" Then
            Response.Write vbcrlf & "<tr" & table3 & "><td colspan=2 align=center height=150>"

            If post_chk() = "no" Then
                Response.Write web_var(web_error,1)
            Else
                red_3         = ""
                accept_u      = Trim(Request.form("accept_u"))
                topic         = Trim(Request.form("topic"))
                word          = Request.form("word")

                If symbol_name(accept_u) <> "yes" Then
                    red_3     = red_3 & "<br><li><font class=red_3>�� �� ��</font> Ϊ�ջ򲻷�����ع���"
                Else
                    sql       = "select username from user_data where username='" & accept_u & "'"
                    Set rs    = conn.execute(sql)

                    If rs.eof And rs.bof Then
                        red_3 = red_3 & "<br><li>����д�� <font class=red_3>�� �� ��</font> ���񲻴��ڣ�"
                    End If

                    rs.Close
                End If

                If var_null(topic) = "" Or Len(topic) > 20 Then
                    red_3 = red_3 & "<br><li><font class=red_3>��������</font> ����Ϊ���ҳ��Ȳ��ܴ���20��"
                End If

                If var_null(word) = "" Or Len(word) > 250 Then
                    red_3 = red_3 & "<br><li><font class=red_3>��������</font> ����Ϊ���ҳ��Ȳ��ܴ���250��"
                End If

                If red_3 = "" Then
                    Set rs = Server.CreateObject("adodb.recordset")
                    sql    = "select * from user_mail"
                    rs.open sql,conn,1,3
                    rs.addnew
                    rs("send_u")     = login_username
                    rs("accept_u")     = accept_u
                    rs("topic")     = topic
                    rs("word")     = word
                    rs("tim")     = now_time

                    If Trim(Request.form("send_later")) = "yes" Then
                        rs("types") = 2
                    Else
                        rs("types") = 1
                    End If

                    rs("isread")     = False
                    rs.update
                    rs.Close

                    If Trim(Request.form("send_later")) = "yes" Then
                        Response.Write "<font class=red>���ѳɹ��ı�����һ�����ŵ����ݣ�</font><br><br><a href='user_mail.asp?action=outbox'>�������</a>"
                    Else
                        Response.Write "<font class=red>���ѳɹ��ĸ� <font class=blue_1><b>" & accept_u & "</b></font> ת����һ�����ţ�</font><br><br><a href='user_mail.asp'>�������</a>"
                    End If

                Else
                    Response.Write found_error(red_3,"250")
                End If

            End If

            Response.Write vbcrlf & "</td></tr>"
        Else
            sql    = "select send_u,topic,word,tim from user_mail where (send_u='" & login_username & "' or accept_u='" & login_username & "') and id=" & id
            Set rs = conn.execute(sql)

            If rs.eof And rs.bof Then
                rs.Close
                red_3 = "<br><li>����ת���� <font class=red_3>����ID</font> �����ڻ��д���"
                red_3 = found_error(red_3,"240")
                Response.Write vbcrlf & "<tr" & table3 & "><td align=center height=150 colspan=2>" & red_3 & "</td></tr>"

                Exit Sub
                Else
                    Response.Write vbcrlf & "<form name=mail_frm action='user_message.asp?action=fw&id=" & id & "' method=post onsubmit=""frm_submitonce(this);""><input type=hidden name=fw_ok value='ok'><input type=hidden name=send_later value=''>" & _
                    vbcrlf & "<tr height=30" & table3 & "><td width='15%' align=center>�� �� �ˣ�</td><td width='85%'>&nbsp;<input type=text name=accept_u size=30 maxlength=20>" & redx & "&nbsp;��&nbsp;" & friend_select() & "</td></tr>" & _
                    vbcrlf & "<tr height=30" & table3 & "><td align=center>�������⣺</td><td>&nbsp;<input type=text name=topic value='FW:" & rs("topic") & "' size=60 maxlength=20></td></tr>" & _
                    vbcrlf & "<tr height=100" & table3 & "><td align=center class=htd>�������ݣ�<br>" & web_var(web_error,3) & "</td><td>&nbsp;<textarea cols=64 rows=6 name=word title='�����������250���ַ�<br>�� Ctrl+Enter ��ֱ�ӷ���' onkeydown=""javascript:frm_quicksubmit();"">����Ϊ " & login_username & " ת�� " & rs("send_u") & " �� " & rs("tim") & " д�Ķ���" & vbcrlf & "������������������������������������������������������������" & vbcrlf & rs("word") & "</textarea></td></tr>" & _
                    vbcrlf & "<tr" & table3 & "><td colspan=2 height=40 align=center><input type=Submit name=wsubmit value='���Ͷ���'>&nbsp;��&nbsp;<input type=Submit name=send value='�������' onclick=""javascript:mail_send_later();"">&nbsp;��&nbsp;<input type=reset value='�����д'></td></tr></form>"
                End If

                rs.Close
            End If

        End Sub

        Sub mail_edit()
            Response.Write vbcrlf & "<tr" & table2 & "><td colspan=2 align=center><font class=end><b>�༩����</b></font></td></tr>"

            If Trim(Request.form("edit_ok")) = "ok" Then
                Response.Write vbcrlf & "<tr" & table3 & "><td colspan=2 align=center height=150>"

                If post_chk() = "no" Then
                    Response.Write web_var(web_error,1)
                Else
                    red_3         = ""
                    accept_u      = Trim(Request.form("accept_u"))
                    topic         = Trim(Request.form("topic"))
                    word          = Request.form("word")

                    If symbol_name(accept_u) <> "yes" Then
                        red_3     = red_3 & "<br><li><font class=red_3>�� �� ��</font> Ϊ�ջ򲻷�����ع���"
                    Else
                        sql       = "select username from user_data where username='" & accept_u & "'"
                        Set rs    = conn.execute(sql)

                        If rs.eof And rs.bof Then
                            red_3 = red_3 & "<br><li>����д�� <font class=red_3>�� �� ��</font> ���񲻴��ڣ�"
                        End If

                        rs.Close
                    End If

                    If var_null(topic) = "" Or Len(topic) > 20 Then
                        red_3 = red_3 & "<br><li><font class=red_3>��������</font> ����Ϊ���ҳ��Ȳ��ܴ���20��"
                    End If

                    If var_null(word) = "" Or Len(word) > 250 Then
                        red_3 = red_3 & "<br><li><font class=red_3>��������</font> ����Ϊ���ҳ��Ȳ��ܴ���250��"
                    End If

                    If red_3 = "" Then
                        Set rs = Server.CreateObject("adodb.recordset")
                        sql    = "select * from user_mail where id=" & id
                        rs.open sql,conn,1,3

                        If rs.eof And rs.bof Then
                            rs.Close:Set rs = Nothing
                            Call close_conn()
                            Call cookies_type(mail_id)
                            Response.End
                        End If

                        rs("send_u")     = login_username
                        rs("accept_u")     = accept_u
                        rs("topic")     = topic
                        rs("word")     = word
                        rs("tim")     = now_time
                        rs("types")     = 1

                        If Trim(Request.form("send_later")) = "yes" Then
                            rs("types") = 2
                        Else
                            rs("types") = 1
                        End If

                        rs("isread")     = False
                        rs.update
                        rs.Close

                        If Trim(Request.form("send_later")) = "yes" Then
                            Response.Write "<font class=red>���ѳɹ��ı����˶��ŵ����ݣ�</font><br><br><a href='user_mail.asp?action=outbox'>�������</a>"
                        Else
                            Response.Write "<font class=red>���ѳɹ��ĸ� <font class=blue_1><b>" & accept_u & "</b></font> ������һ�����ţ�</font><br><br><a href='user_mail.asp'>�������</a>"
                        End If

                    Else
                        Response.Write found_error(red_3,"250")
                    End If

                End If

                Response.Write vbcrlf & "</td></tr>"
            Else
                sql    = "select accept_u,topic,word,tim from user_mail where (send_u='" & login_username & "' or accept_u='" & login_username & "') and id=" & id
                Set rs = conn.execute(sql)

                If rs.eof And rs.bof Then
                    rs.Close
                    red_3 = "<br><li>�����༩�� <font class=red_3>����ID</font> �����ڻ��д���"
                    red_3 = found_error(red_3,"240")
                    Response.Write vbcrlf & "<tr><td align=center height=150 colspan=2>" & red_3 & "</td></tr>"

                    Exit Sub
                    Else
                        Response.Write vbcrlf & "<form name=mail_frm action=user_message.asp?action=edit&id=" & id & " method=post onsubmit=""frm_submitonce(this);""><input type=hidden name=edit_ok value='ok'><input type=hidden name=send_later value=''>" & _
                        vbcrlf & "<tr height=30" & table3 & "><td width='15%' align=center>�� �� �ˣ�</td><td width='85%'>&nbsp;<input type=text name=accept_u value='" & rs("accept_u") & "' size=30 maxlength=20>" & redx & "&nbsp;��&nbsp;" & friend_select() & "</td></tr>" & _
                        vbcrlf & "<tr height=30" & table3 & "><td align=center>�������⣺</td><td>&nbsp;<input type=text name=topic value='" & rs("topic") & "' size=60 maxlength=20></td></tr>" & _
                        vbcrlf & "<tr height=100" & table3 & "><td align=center class=htd>�������ݣ�<br>" & web_var(web_error,3) & "</td><td>&nbsp;<textarea cols=64 rows=6 name=word title='�����������250���ַ�<br>�� Ctrl+Enter ��ֱ�ӷ���' onkeydown=""javascript:frm_quicksubmit();"">" & rs("word") & "</textarea></td></tr>" & _
                        vbcrlf & "<tr" & table3 & "><td colspan=2 height=40 align=center><input type=Submit name=wsubmit value='���Ͷ���'>&nbsp;��&nbsp;<input type=Submit name=send value='�������' onclick=""javascript:mail_send_later();"">&nbsp;��&nbsp;<input type=reset value='�����д'></td></tr></form>"
                    End If

                    rs.Close
                End If

            End Sub

            Function mail_view()
                mail_view = vbcrlf & "<tr" & table2 & " height=25><td align=center background=images/" & web_var(web_config,5) & "/bar_3_bg.gif><font class=end><b>�鿴����</b></font></td></tr>"
                red_3     = ""
                sql       = "select * from user_mail where (send_u='" & login_username & "' or accept_u='" & login_username & "') and id=" & id
                Set rs    = conn.execute(sql)

                If rs.eof And rs.bof Then
                    rs.Close:Set rs = Nothing
                    red_3     = "<br><li>�����鿴�� <font class=red_3>����ID</font> �����ڻ��д���"
                    red_3     = found_error(red_3,"240")
                    mail_view = mail_view & "<tr" & table3 & "><td align=center height=150>" & red_3 & "</td></tr>"
                    Exit Function
                End If

                send_u    = rs("send_u")
                accept_u  = rs("accept_u")
                types     = Int(rs("types"))
                isread    = rs("isread")
                mail_view = mail_view & vbcrlf & "<tr" & table3 & "><td height=50>&nbsp;&nbsp;�������⣺<font class=red_3>" & code_html(rs("topic"),1,0) & "</font></td></tr>" & _
                vbcrlf & "<tr" & table3 & "><td height=80 align=center><table border=0 width='96%' class=tf><tr><td height=8></td></tr><tr><td class=bw>" & code_jk(rs("word")) & "</td></tr><tr><td height=8></td></tr></table></td></tr>" & _
                vbcrlf & "<tr" & table3 & "><td align=center height=30>������ " & format_user_view(send_u,1,1) & " �� " & time_type(rs("tim"),88) & " �������͵Ķ���</td></tr>"
                rs.Close:Set rs = Nothing

                If Not(send_u = login_username And accept_u <> login_username) And isread = False Then
                    sql = "update user_mail set isread=1 where types<>2 and id=" & id
                    conn.execute(sql)
                    If login_message > 0 Then login_message = login_message - 1
                End If

            End Function

            Function friend_select()
                Dim sql
                Dim rs
                Dim ttt
                friend_select = vbcrlf & "<script language=javascript>" & _
                vbcrlf & "function Do_accept(addaccept) {" & _
                vbcrlf & "  if (addaccept!=0) { document.mail_frm.accept_u.value=addaccept; }" & _
                vbcrlf & "  return;" & _
                vbcrlf & "}</script>" & _
                vbcrlf & "<select name=friend_select size=1 onchange=Do_accept(this.options[this.selectedIndex].value)>" & _
                vbcrlf & "<option value='0'>ѡ���ҵĺ���</option>"
                sql               = "select username2 from user_friend where username1='" & login_username & "' order by id"
                Set rs            = conn.execute(sql)

                Do While Not rs.eof
                    ttt           = rs(0)
                    friend_select = friend_select & vbcrlf & "<option value='" & ttt & "'>" & ttt & "</option>"
                    rs.movenext
                Loop

                rs.Close
                friend_select = friend_select & vbcrlf & "</select>"
            End Function %>
<script language=javascript>
<!--
//���÷���:onsubmit="frm_submitonce(this);"
function frm_submitonce(theform)
{
  if (document.all||document.getElementById)
  {
    for (i=0;i<theform.length;i++)
    {
      var tempobj=theform.elements[i]
      if(tempobj.type.toLowerCase()=="submit"||tempobj.type.toLowerCase()=="reset")
      tempobj.disabled=true
    }
  }
}

function frm_quicksubmit(eventobject)
{
  if (event.keyCode==13 && event.ctrlKey)
  mail_frm.wsubmit.click();
}

function mail_send_later()
{
  this.document.mail_frm.send_later.value='yes';
}
-->
</script>
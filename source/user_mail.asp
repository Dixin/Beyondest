<!-- #include file="INCLUDE/config_user.asp" -->
<!-- #include file="include/conn.asp" -->
<%
' ====================
' Beyondest.Com v3.6.1
' http://beyondest.com
' ====================

Dim nummer,rssum,action_temp
tit    = "վ�ڶ���"
nummer = 0

Call web_head(2,0,0,0,0)
'------------------------------------left----------------------------------
Call left_user()
'----------------------------------left end--------------------------------
Call web_center(0)
'-----------------------------------center---------------------------------
Response.Write ukong
Call user_mail_menu(0)
Response.Write table1 %>
<tr align=center<% Response.Write table2 %> height=25>
<td width='6%' background=images/<% = web_var(web_config,5) %>/bar_3_bg.gif><font class=end><b>�Ѷ�</b></font></td>
<td width='20%' background=images/<% = web_var(web_config,5) %>/bar_3_bg.gif><font class=end><b><%

If action = "outbox" Or action = "issend" Then
    Response.Write "��"
Else
    Response.Write "��"
End If %>����</b></font></td>
<td width='38%' background=images/<% = web_var(web_config,5) %>/bar_3_bg.gif><font class=end><b>��������</b></font></td>
<td width='20%' background=images/<% = web_var(web_config,5) %>/bar_3_bg.gif><font class=end><b>��������</b></font></td>
<td width='10%' background=images/<% = web_var(web_config,5) %>/bar_3_bg.gif><font class=end><b>��С</b></font></td>
<td width='6%' background=images/<% = web_var(web_config,5) %>/bar_3_bg.gif><font class=end><b>����</b></font></td>
</tr>
<script language=javascript src='STYLE/admin_del.js'></script>
<form name=del_form action='user_mail.asp?action=<% Response.Write action %>' method=post>
<input type=hidden name=action2 value='delete'>
<input type=hidden name=del_type value='<% Response.Write action %>'>
<%

If Trim(Request.form("action2")) = "delete" And Len(Trim(Request.form("del_sel"))) Then
    Response.Write del_select()
End If

Function del_select()
    Dim delid,del_i,del_num,del_dim,del_sql,del_type
    del_type = Trim(Request.form("del_type"))
    delid    = Trim(Request.form("del_id"))

    Select Case del_type
        Case "outbox","issend"
            del_sql = "update user_mail set types=4 where send_u='" & login_username & "' and id="
        Case "recycle"
            del_sql = "delete from user_mail where (send_u='" & login_username & "' or accept_u='" & login_username & "') and id="
        Case Else
            del_sql = "update user_mail set types=4 where accept_u='" & login_username & "' and id="
    End Select

    If var_null(delid) <> "" Then
        delid   = Replace(delid," ","")
        del_dim = Split(delid,",")
        del_num = UBound(del_dim)

        For del_i = 0 To del_num
            conn.execute(del_sql & del_dim(del_i))
        Next

        Erase del_dim

        If del_type = "recycle" Then
            del_select = "����ɾ���ɹ�����ɾ���� " & del_num + 1 & " �����š�\n\n�����ѳ���ɾ����"
        Else
            del_select = "����ɾ���ɹ�����ɾ���� " & del_num + 1 & " �����š�\n\nɾ���Ķ��Ž��������Ļ���վ�ڡ�"
        End If

        del_select     = vbcrlf & "<script language=javascript>alert(""" & del_select & """);</script>"
    End If

End Function

If Len(Trim(Request.form("clear"))) > 0 Then
    Response.Write mail_clear()
End If

sql = "select * from user_mail where "

Select Case action
    Case "outbox"
        sql         = sql & "send_u='" & login_username & "' and types=2"
        action_temp = "�ݸ���"
    Case "issend"
        sql         = sql & "send_u='" & login_username & "' and types=1"
        action_temp = "�ѷ�����"
    Case "recycle"
        sql         = sql & "(accept_u='" & login_username & "' or send_u='" & login_username & "') and types=4"
        action_temp = "������"
    Case Else
        action      = "inbox"
        sql         = sql & "accept_u='" & login_username & "' and types=1"
        action_temp = "������"
End Select

sql           = sql & " order by id desc"
login_message = 0
Set rs        = Server.CreateObject("adodb.recordset")
rs.open sql,conn,1,1

If Not(rs.eof And rs.bof) Then
    rssum  = rs.recordcount
    nummer = rssum

    For i = 1 To rssum
        Response.Write mail_type(rs)
        rs.movenext
    Next

End If %>
<tr><td colspan=6 bgcolor=<% Response.Write web_var(web_color,5) %> height=30 align=center class=htd>��<font class=red><% Response.Write nummer %></font>������<font class=gray>��Ϊ�˽�ʡ�ռ䣬�뼰ʱɾ��������Ϣ��</font>
<input type=checkbox name=del_all value=1 onClick=selectall('<% Response.Write nummer %>') class=bg_3> ѡ������
<input type=submit name=del_sel value='ɾ����ѡ' onclick="return suredel('<% Response.Write nummer %>');">
<input type=submit name=clear onclick="{if(confirm('ȷ�����<% Response.Write action_temp %>���еļ�¼��?\n\n��պ��޷��ָ���')){this.document.del_form.submit();return true;}return false;}" value="���<% Response.Write action_temp %>" style='width:90px'></td></tr>
</table>
<%
Response.Write ukong
'---------------------------------center end-------------------------------
Call web_end(0)

Function mail_clear()
    Dim clear_type

    Select Case Trim(Request.form("del_type"))
        Case "inbox"
            conn.execute("delete from user_mail where accept_u='" & login_username & "' and types=1")
            clear_type = "������"
        Case "outbox"
            conn.execute("delete from user_mail where send_u='" & login_username & "' and types=2")
            clear_type = "�ݸ���"
        Case "issend"
            conn.execute("delete from user_mail where send_u='" & login_username & "' and types=1")
            clear_type = "�ѷ�����"
        Case "recycle"
            conn.execute("delete from user_mail where (accept_u='" & login_username & "' or send_u='" & login_username & "') and types=4")
            clear_type = "������"
    End Select

End Function

Function mail_type(rs)
    Dim ttim,isread,td_temp,read_pic,iid,link_temp,name_temp
    td_temp               = ""
    read_pic              = "olds"
    link_temp             = "view"
    iid                   = rs("id"):isread = rs("isread"):ttim = rs("tim")

    If isread = False Then
        td_temp           = " class=btd"
        read_pic          = "news"

        If action = "inbox" Then
            login_message = login_message + 1
        End If

    End If

    If action = "outbox" Then
        td_temp   = " class=btd"
        read_pic  = "sends"
        link_temp = "edit"
    End If

    If action = "outbox" Or action = "issend" Then
        name_temp = format_user_view(rs("accept_u"),1,1)
    Else
        name_temp = format_user_view(rs("send_u"),1,1)
    End If

    ttim = time_type(ttim,8)
    mail_type = vbcrlf & "<tr align=center" & td_temp & table3 & "><td><img src='images/mail/" & read_pic & ".gif' border=0></td><td>" & name_temp & "</td><td align=left><a href='user_message.asp?action=" & link_temp & "&id=" & iid & "'>" & cuted(rs("topic"),15) & "</a></td><td class=timtd>" & ttim & "</td><td>" & Len(rs("word")) & "Byte</td><td><input type=checkbox name=del_id value='" & iid & "' class=bg_1></td></tr>"
End Function %>
<!-- #include file="include/onlogin.asp" -->
<!-- #include file="include/conn.asp" -->
<%
' ====================
' Beyondest.Com v3.6.1
' http://beyondest.com
' ====================

Dim rssnum,j,id,vid,vname,nid
tit = "<a href='?'>�鿴���е����б�</a> �� <a href='?action=add'>����µ����б�</a>"
Response.Write header(8,tit)
id  = Trim(Request.querystring("id"))
vid = Trim(Request.querystring("vid"))

Select Case action
    Case "add"
        Call vote_add()
    Case "edit"
        Call vote_edit()
    Case "edit2"
        Call vote_edit2()
    Case "view"
        Call vote_view()
    Case "del"
        Call vote_del()
    Case "delete"
        Call vote_delete()
    Case Else
        Call vote_main()
End Select

Call close_conn()
Response.Write ender()

Sub vote_del()

    If Not(IsNumeric(id)) Then Call vote_main():Exit Sub
        conn.execute("delete from vote where vtype=1 and id=" & id)
        Response.Write "<script language=javascript>alert(""�ɹ�ɾ���˵�����Ŀ��" & id & "����\n\n������ء���"");location.href='?action=view&vid=" & vid & "';</script>"
    End Sub

    Sub vote_delete()

        If Not(IsNumeric(vid)) Then Call vote_main():Exit Sub
            conn.execute("delete from vote where vid=" & vid)
            Response.Write "<script language=javascript>alert(""�ɹ�ɾ���˵����б�" & vid & "����\n\n������ء���"");location.href='?';</script>"
        End Sub

        Sub vote_edit2()

            If Not(IsNumeric(id)) Then Call vote_main():Exit Sub
                sql    = "select vid,vname,counter from vote where vtype=1 and id=" & id
                Set rs = conn.execute(sql)

                If rs.eof And rs.bof Then
                    rs.Close:Set rs = Nothing
                    Response.Write "<script language=javascript>alert(""������Ŀ�����ڣ�\n\n������ء���"");location.href='?';</script>"

                    Exit Sub
                    End If

                    Dim counter
                    vid = rs("vid"):vname = rs("vname"):counter = rs("counter")
                    rs.Close:Set rs = Nothing

                    If Trim(Request.querystring("chk")) = "yes" Then
                        counter = code_admin(Request.form("counter"))
                        If Not(IsNumeric(counter)) Then counter =  - 1
                        vname   = code_admin(Request.form("vname"))

                        If Int(counter) < 0 Or InStr(1,counter,".") > 0 Then

                            Response.Write "<font class=red_2>ͶƱ����ֻ��Ϊ�������Ҳ���Ϊ�գ�</font><br><br>" & go_back:Exit Sub
                            End If

                            If Len(vname) < 1 Then

                                Response.Write "<font class=red_2>��Ŀ���Ʋ���Ϊ�գ�</font><br><br>" & go_back:Exit Sub
                                End If

                                sql = "update vote set vname='" & vname & "',counter=" & counter & " where vtype=1 and id=" & id
                                conn.execute(sql)
                                Response.Write "<script language=javascript>alert(""�ɹ��޸���һ��������Ŀ���ƣ�\n\n������ء���"");location.href='?action=view&vid=" & vid & "';</script>"

                                Exit Sub
                                End If %>
<table border=0>
<form action='?action=edit2&id=<% Response.Write id %>&chk=yes' method=post>
<tr><td colspan=2 align=center height=50><a href='?action=view&vid=<% Response.Write vid %>' class=red>�޸����е�����Ŀ</a></td></tr>
<tr><td>��Ŀ���ƣ�</td><td><input type=text name=vname value='<% Response.Write vname %>' size=30 maxlength=20></td></tr>
<tr><td height=30>ͶƱ������</td><td><input type=text name=counter value='<% Response.Write counter %>' size=10 maxlength=10><% Response.Write redx %>ֻ��Ϊ0��������</td></tr>
<tr><td colspan=2 align=center><input type=submit value='�� �� �� ��'>����<input type=reset value='������д'></td></tr>
</form>
</table>
<%
                            End Sub

                            Sub vote_edit()

                                If Not(IsNumeric(vid)) Then Call vote_main():Exit Sub
                                    sql    = "select id,vname from vote where vtype=0 and vid=" & vid
                                    Set rs = conn.execute(sql)

                                    If rs.eof And rs.bof Then
                                        rs.Close:Set rs = Nothing
                                        Response.Write "<script language=javascript>alert(""�����б�" & vid & "�������ڣ�\n\n������ء���"");location.href='?';</script>"

                                        Exit Sub
                                        End If

                                        vname = rs("vname")
                                        rs.Close:Set rs = Nothing

                                        If Trim(Request.querystring("chk")) = "yes" Then
                                            vname = code_admin(Request.form("vname"))

                                            If Len(vname) < 1 Then

                                                Response.Write "<font class=red_2>�������Ʋ���Ϊ�գ�</font><br><br>" & go_back:Exit Sub
                                                End If

                                                sql = "update vote set vname='" & vname & "' where vtype=0 and vid=" & vid
                                                conn.execute(sql)
                                                Response.Write "<script language=javascript>alert(""�ɹ��޸��˵����б�" & vid & "�������ƣ�\n\n������ء���"");location.href='?action=view&vid=" & vid & "';</script>"

                                                Exit Sub
                                                End If %>
<table border=0>
<form action='?action=edit&vid=<% Response.Write vid %>&chk=yes' method=post>
<tr><td colspan=2 align=center height=50 class=red>�޸ĵ����б�����</td></tr>
<tr><td>���� ID��</td><td><input type=text name=vid value='<% Response.Write vid %>' size=10 maxlength=10 disabled><% Response.Write redx %>ֻ��Ϊ������</td></tr>
<tr><td height=50>�������ƣ�</td><td><input type=text name=vname value='<% Response.Write vname %>' size=30 maxlength=20><% Response.Write redx %></td></tr>
<tr><td colspan=2 align=center><input type=submit value='�� �� �� ��'>����<input type=reset value='������д'></td></tr>
</form>
</table>
<%
                                            End Sub

                                            Sub vote_add()

                                                If Trim(Request.querystring("chk")) = "yes" Then
                                                    vid   = code_admin(Request.form("vid"))
                                                    If Not(IsNumeric(vid)) Then vid = 0
                                                    vname = code_admin(Request.form("vname"))

                                                    If Int(vid) < 1 Or InStr(1,vid,".") > 0 Then

                                                        Response.Write "<font class=red_2>�����б� ID ֻ��Ϊ�������Ҳ���Ϊ�գ�</font><br><br>" & go_back:Exit Sub
                                                        End If

                                                        If Len(vname) < 1 Then

                                                            Response.Write "<font class=red_2>�������Ʋ���Ϊ�գ�</font><br><br>" & go_back:Exit Sub
                                                            End If

                                                            sql    = "select id from vote where vtype=0 and vid=" & vid
                                                            Set rs = conn.execute(sql)

                                                            If Not(rs.eof And rs.bof) Then
                                                                rs.Close:Set rs = Nothing

                                                                Response.Write "<font class=red_2>�����б� ID��" & vid & "���Ѵ��ڣ����������롣</font><br><br>" & go_back:Exit Sub
                                                                End If

                                                                rs.Close:Set rs = Nothing
                                                                sql = "insert into vote(vid,vtype,vname,counter) values(" & vid & ",0,'" & vname & "',0)"
                                                                conn.execute(sql)
                                                                Response.Write "<script language=javascript>alert(""�ɹ������һ���µĵ����б�\n\n������ء���"");location.href='?';</script>"

                                                                Exit Sub
                                                                End If %>
<table border=0>
<form action='?action=add&chk=yes' method=post>
<tr><td colspan=2 align=center height=50 class=red>����µĵ����б�</td></tr>
<tr><td>���� ID��</td><td><input type=text name=vid size=10 maxlength=10><% Response.Write redx %>ֻ��Ϊ������</td></tr>
<tr><td height=50>�������ƣ�</td><td><input type=text name=vname size=30 maxlength=20><% Response.Write redx %></td></tr>
<tr><td colspan=2 align=center><input type=submit value='�� �� �� ��'>����<input type=reset value='������д'></td></tr>
</form>
</table>
<%
                                                            End Sub

                                                            Sub vote_view()

                                                                If Not(IsNumeric(vid)) Then Call vote_main():Exit Sub

                                                                    If Trim(Request.querystring("chk")) = "yes" Then
                                                                        vname = code_admin(Request.form("vname"))

                                                                        If Len(vname) < 1 Then

                                                                            Response.Write "<font class=red_2>������Ŀ����Ϊ�գ�</font><br><br>" & go_back:Exit Sub
                                                                            End If

                                                                            sql = "insert into vote(vid,vtype,vname,counter) values(" & vid & ",1,'" & vname & "',0)"
                                                                            conn.execute(sql)
                                                                            Response.Write "<script language=javascript>alert(""�ɹ������һ���µ�����Ŀ��\n\n������ء���"");location.href='?action=view&vid=" & vid & "';</script>"

                                                                            Exit Sub
                                                                            End If %>
<table border=1 width=400 cellspacing=0 cellpadding=2<% Response.Write table1 %>>
<%
                                                                            sql    = "select id,vid,vname,counter from vote where vid=" & vid & " order by id"
                                                                            Set rs = conn.execute(sql)

                                                                            If rs.eof And rs.bof Then
                                                                                rs.Close:Set rs = Nothing
                                                                                Response.Write "<script language=javascript>alert(""�����б�" & vid & "�������ڣ�\n\n������ء���"");location.href='?';</script>"

                                                                                Exit Sub
                                                                                End If

                                                                                j       = 0

                                                                                Do While Not rs.eof
                                                                                    nid = rs("id")

                                                                                    If j = 0 Then %>
<tr>
<td colspan=2 height=25 bgcolor=<% Response.Write color3 %> class=red_3>&nbsp;&nbsp;<b><% Response.Write code_html(rs("vname"),1,0) %></b>��ID��<% Response.Write vid %>��</td>
<td align=center><a href='?action=edit&vid=<% Response.Write vid %>'>�༭����</a></td>
</td></tr>
<% Else %>
<tr align=center<% Response.Write mtr %>>
<td width='8%'><% Response.Write j %></td>
<td width='76%' align=left><% Response.Write rs("vname") %> <font class=blue><% Response.Write rs("counter") %></font></td>
<td width='16%'><a href='?action=edit2&id=<% Response.Write nid %>'>�༭</a> <a href="javascript:do_del(<% Response.Write vid %>,<% Response.Write nid %>);">ɾ��</a></td>
</tr>
<%
                                                                                    End If

                                                                                    j = j + 1
                                                                                    rs.movenext
                                                                                Loop

                                                                                rs.Close:Set rs = Nothing %>
<tr><td colspan=3 height=25 align=center>
  <table border=0>
  <form action='?action=view&vid=<% Response.Write vid %>&chk=yes' method=post>
  <tr>
  <td>�µ���Ŀ���ƣ�</td>
  <td><input type=text name=vname size=20 maxlength=20></td>
  <td>&nbsp;&nbsp;<input type=submit value='������'></td>
  </tr>
  </form>
  </table>
</td></tr>
</table>
<%
                                                                            End Sub

                                                                            Sub vote_main() %>
<table border=1 width=400 cellspacing=0 cellpadding=2<% Response.Write table1 %>>
<tr align=center height=20 bgcolor=<% Response.Write color3 %>>
<td width='8%'>ID</td>
<td width='76%'>�����б�����</td>
<td width='16%'>����</td>
</tr>
<%
                                                                                sql = "select id,vid,vname from vote where vtype=0 order by id desc"
                                                                                Set rs = conn.execute(sql)

                                                                                Do While Not rs.eof
                                                                                    nid = rs("id"):vid = rs("vid") %>
<tr align=center<% Response.Write mtr %>>
<td class=blue><b><% Response.Write vid %></b></td>
<td align=left><a href='?action=view&vid=<% Response.Write vid %>'><% Response.Write code_html(rs("vname"),1,0) %></a></td>
<td><a href='?action=edit&vid=<% Response.Write vid %>'>�༭</a> <a href="javascript:do_delete(<% Response.Write vid %>);">ɾ��</a></td>
</tr>
<%
                                                                                    rs.movenext
                                                                                Loop

                                                                                rs.Close:Set rs = Nothing %>
</table>
<br>
<table border=0 width=450>
<tr><td colspan=2>���÷�����</td></tr>
<tr><td colspan=2 height=40>&lt;script language=javascript src='vote.asp?id=<font class=red>1</font>&types=<font class=red>1</font>&mcolor=<font class=red>ff0000</font>&bgcolor=<font class=red>ededed</font>'&gt;&lt;/script&gt;</td></tr>
<tr><td>ʹ��˵����</td><td>1����һ��������Ҫ���õĵ���ID��</td></tr>
<tr><td></td><td>2���ڶ��������ǵ�����ʾ�����ͣ���1��Ϊ��ѡ����2��Ϊ��ѡ��</td></tr>
<tr><td></td><td>3�������������ǵ��������ʾ��ɫ������Ҫ�ӡ�#����</td></tr>
<tr><td></td><td>4�����ĸ������ǵ���ѡ��򱳾�ɫ������Ҫ�ӡ�#����</td></tr>
</table>
<%
                                                                            End Sub %>
<script language=JavaScript><!--
function do_del(data1,data2)
{
  if (confirm("�˲�����ɾ��IDΪ "+data2+" �ĵ�����Ŀ��\n\n���Ҫɾ����\nɾ�����޷��ָ���"))
    window.location="?action=del&vid="+data1+"&id="+data2
}
function do_delete(data1)
{
  if (confirm("�˲�����ɾ��IDΪ "+data1+" �ĵ����б�\n\n���Ҫɾ����\nɾ�����޷��ָ���"))
    window.location="?action=delete&vid="+data1
}
//--></script>
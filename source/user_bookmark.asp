<!-- #include file="include/config_user.asp" -->
<!-- #include file="include/conn.asp" -->
<%
' ====================
' Beyondest.Com v3.6.1
' http://beyondest.com
' ====================

Dim n_sort
Dim gid
Dim g_id
Dim gname
Dim select_add
Dim id
Dim name
Dim url
Dim rssum
tit  = "������ǩ":n_sort = "book":rssum = 0
g_id = Trim(Request.querystring("g_id"))
If Not(IsNumeric(g_id)) Then g_id = 0

Call web_head(2,0,0,0,0)
'------------------------------------left----------------------------------
Call left_user()
'----------------------------------left end--------------------------------
Call web_center(0)
'-----------------------------------center---------------------------------

Select Case action
    Case "groupedit"
        Call group_edit()
    Case "bookmarkedit"
        Call bookmark_edit()
    Case "groupdel"
        Call group_del()
    Case "bookmarkdel"
        Call bookmark_del()
    Case "groupadd"
        Call group_add()
    Case "bookmarkadd"
        Call bookmark_add()
End Select %>
<% Response.Write ukong & table1 %>
<tr<% Response.Write table2 %> height=25>
<td class=end width='90%' background=images/<% = web_var(web_config,5) %>/bar_3_bg.gif>&nbsp;<% Response.Write img_small(us) %>&nbsp;<b>�ҵ���ǩ��</b></td>
<td class=end width='10%' align=center background=images/<% = web_var(web_config,5) %>/bar_3_bg.gif><b>�����</b></td>
</tr>
<form name=groupedit_frm action='?action=groupedit' method=post>
<input type=hidden name=g_id value=''>
<input type=hidden name=g_name value=''>
</form>
<tr<% Response.Write table3 %>>
<td><% Response.Write img_small("jt0") %><a href='?g_id=0' class=gray>[ ����ǩ�� ]</a>&nbsp;&nbsp;
<% Response.Write img_small("jt0") %><a href='?action=all' class=gray>[ ���������ǩ ]</a></td>
<td align=center class=gray>��</td>
</tr>
<%
sql            = "select g_id,g_name from jk_group where g_sort='" & n_sort & "' and username='" & login_username & "' order by g_id"
Set rs         = conn.execute(sql)

Do While Not rs.eof
    gid        = rs("g_id"):gname = rs("g_name")
    select_add = select_add & vbcrlf & "<option value='" & gid & "'"
    If Int(gid) = Int(g_id) Then select_add = select_add & " selected"
    select_add = select_add & ">" & gname & "</option>" %>
<tr<% Response.Write table3 %>>
<td><% Response.Write img_small("jt0") %><a href='?g_id=<% Response.Write gid %>'<% If Int(g_id) = Int(gid) Then Response.Write " class=red_3" %>><% Response.Write gname %></a></td>
<td align=center><a href="javascript:group_edit(<% Response.Write gid %>,'<% Response.Write gname %>');"><img src='IMAGES/SMALL/EDIT.GIF' border=0 title='�޸�'></a>&nbsp;<a href="javascript:group_del(<% Response.Write gid %>);"><img src='IMAGES/SMALL/DEL.GIF' border=0 title='ɾ��'></a></td>
</tr>
<%
    rs.movenext
Loop

rs.Close %>
</table>
<% Response.Write kong & table1 %>
<tr<% Response.Write table2 %> align=center height=25>
<td class=end width='6%' background=images/<% = web_var(web_config,5) %>/bar_3_bg.gif><b>���</b></td>
<td class=end width='34%' background=images/<% = web_var(web_config,5) %>/bar_3_bg.gif><b>�ҵĸ�����ǩ����</b></td>
<td class=end width='50%' background=images/<% = web_var(web_config,5) %>/bar_3_bg.gif><b>��ǩ��ַ</b></td>
<td class=end width='10%' background=images/<% = web_var(web_config,5) %>/bar_3_bg.gif><b>�� ��</b></td>
</tr>
<form name=bookmarkedit_frm action='?action=bookmarkedit' method=post>
<input type=hidden name=id value=''>
<input type=hidden name=name value=''>
<input type=hidden name=url value=''>
</form>
<%
sql       = "select id,name,url from user_bookmark where"

If action <> "all" Then
    sql   = sql & " g_id=" & g_id & " and"
End If

sql       = sql & " username='" & login_username & "' order by id desc"
Set rs    = conn.execute(sql)

Do While Not rs.eof
    id    = rs("id"):name = rs("name"):url = rs("url")
    rssum = rssum + 1 %>
<tr<% Response.Write table3 %>>
<td align=center><% Response.Write rssum %></td>
<td><a href='<% Response.Write url %>' target=_blank title='<% Response.Write code_html(name,1,0) %>'><% Response.Write code_html(name,1,15) %></a></td>
<td><a href='<% Response.Write url %>' target=_blank title='<% Response.Write code_html(url,1,0) %>'><% Response.Write code_html(url,1,25) %></a></td>
<td align=center><a href="javascript:bookmark_edit(<% Response.Write id %>,'<% Response.Write name %>','<% Response.Write url %>');"><img src='IMAGES/SMALL/EDIT.GIF' border=0></a>&nbsp;<a href="javascript:bookmark_del(<% Response.Write id %>);"><img src='IMAGES/SMALL/DEL.GIF' border=0></a></td>
</tr>
<%
    rs.movenext
Loop

rs.Close %>
</table>
<% Response.Write kong & table1 %>
<tr<% Response.Write table2 %> height=25>
<td class=end background=images/<% = web_var(web_config,5) %>/bar_3_bg.gif>&nbsp;<% Response.Write img_small(us) %>&nbsp;<b>����µ���ǩ��</b></td>
</tr>
<tr<% Response.Write table3 %>><td>
  <table border=0 cellpadding=5>
  <form action='?action=groupadd' method=post>
  <tr>
  <td>�������ƣ�</td>
  <td><input type=text name=g_name size=20 maxlength=20></td>
  <td><input type=submit value='�����ǩ��'></td>
  </tr>
  </form>
  </table>
</td></tr>
</table>
<% Response.Write kong & table1 %>
<tr<% Response.Write table2 %> height=25>
<td class=end background=images/<% = web_var(web_config,5) %>/bar_3_bg.gif>&nbsp;<% Response.Write img_small(us) %>&nbsp;<b>����µĸ�����ǩ</b></td>
</tr>
<tr<% Response.Write table3 %>><td>
  <table border=0 cellpadding=2>
  <form action='?action=bookmarkadd' method=post>
  <tr>
  <td>����ǩ���ƣ�</td>
  <td>
    <table border=0>
    <tr>
    <td><input type=text name=name size=30 maxlength=50></td>
    <td>����ǩ�飺</td>
    <td><select name=g_id>
    <option value='0'>[ ����ǩ�� ]</option>
<% Response.Write select_add %>
</select></td>
    </tr>
    </table>
  </td></tr>
  <tr>
  <td>����ǩ��ַ��</td>
  <td>
    <table border=0>
    <tr>
    <td><input type=text name=url size=50 value='http://' maxlength=100></td>
    <td>��<input type=submit value='�����ǩ'></td>
    </tr>
    </table>
  </td></tr>
  </form>
  </table>
</td></tr>
</table>
<br>
<%
'---------------------------------center end-------------------------------
Call web_end(0)

Sub group_del()
    gid = Trim(Request.querystring("g_id"))

    If Not(IsNumeric(gid)) Then
        Response.Write("<script language=javascript>" & _
        vbcrlf & "alert(""����ɾ����ǩ��ʱ�������ǩ��ID����\n\n�뷵���������롣"");" & _
        vbcrlf & "history.back(1)" & _
        vbcrlf & "</script>")
        close_conn

        Exit Sub
        End If

        sql = "delete from jk_group where g_id=" & gid & " and g_sort='" & n_sort & "' and username='" & login_username & "'"
        conn.execute(sql)
        sql = "delete from user_bookmark where g_id=" & gid & " and username='" & login_username & "'"
        conn.execute(sql)
        Response.Write("<script language=javascript>" & _
        vbcrlf & "alert(""�ɹ���ɾ����һ��ǩ�飡"");" & _
        vbcrlf & "</script>")
    End Sub

    Sub group_edit()
        gname = code_form(Request.form("g_name"))
        gid   = Trim(Request.form("g_id"))

        If Len(gname) < 1 Or Len(gname) > 20 Or Not(IsNumeric(gid)) Then
            Response.Write("<script language=javascript>" & _
            vbcrlf & "alert(""�����޸���ǩ��� ������ ʱ�������������\n\n�뷵���������롣"");" & _
            vbcrlf & "history.back(1)" & _
            vbcrlf & "</script>")
            close_conn

            Exit Sub
            End If

            sql  = "update jk_group set g_name='" & gname & "' where g_id=" & gid & " and g_sort='" & n_sort & "' and username='" & login_username & "'"
            conn.execute(sql)
            g_id = gid
            Response.Write("<script language=javascript>" & _
            vbcrlf & "alert(""�ɹ����޸�����ǩ������ƣ�" & gname & """);" & _
            vbcrlf & "</script>")
        End Sub

        Sub bookmark_del()
            id = Trim(Request.querystring("id"))

            If Not(IsNumeric(id)) Then
                Response.Write("<script language=javascript>" & _
                vbcrlf & "alert(""����ɾ��������ǩʱ�������ǩID����\n\n�뷵���������롣"");" & _
                vbcrlf & "history.back(1)" & _
                vbcrlf & "</script>")
                close_conn

                Exit Sub
                End If

                sql = "delete from user_bookmark where id=" & id & " and username='" & login_username & "'"
                conn.execute(sql)
                Response.Write("<script language=javascript>" & _
                vbcrlf & "alert(""�ɹ���ɾ����һ��ǩ�飡"");" & _
                vbcrlf & "</script>")
            End Sub

            Sub bookmark_edit()
                name = code_form(Request.form("name"))
                url  = code_form(Request.form("url"))
                id   = Trim(Request.form("id"))

                If Len(name) < 1 Or Len(name) > 50 Or Len(url) < 1 Or Len(url) > 100 Or Not(IsNumeric(id)) Then
                    Response.Write("<script language=javascript>" & _
                    vbcrlf & "alert(""�����޸ĸ�����ǩʱ�������������\n\n�뷵���������롣"");" & _
                    vbcrlf & "history.back(1)" & _
                    vbcrlf & "</script>")
                    close_conn

                    Exit Sub
                    End If

                    sql = "update user_bookmark set name='" & name & "',url='" & url & "' where id=" & id & " and username='" & login_username & "'"
                    conn.execute(sql)
                    Response.Write("<script language=javascript>" & _
                    vbcrlf & "alert(""�ɹ����޸��˸�����ǩ�����ƣ�" & name & "����"");" & _
                    vbcrlf & "</script>")
                End Sub

                Sub group_add()
                    gname = code_form(Request.form("g_name"))

                    If Len(gname) < 1 Or Len(gname) > 20 Then
                        Response.Write("<script language=javascript>" & _
                        vbcrlf & "alert(""�����ǩ��� ������ �Ǳ���Ҫ�ģ�\n\n�뷵�������롣"");" & _
                        vbcrlf & "history.back(1)" & _
                        vbcrlf & "</script>")
                        close_conn

                        Exit Sub
                        End If

                        sql = "insert into jk_group(g_sort,g_name,username) values('" & n_sort & "','" & gname & "','" & login_username & "')"
                        conn.execute(sql)
                        Response.Write("<script language=javascript>" & _
                        vbcrlf & "alert(""�ɹ��������һ��ǩ�飺" & gname & """);" & _
                        vbcrlf & "</script>")
                    End Sub

                    Sub bookmark_add()
                        Dim gg
                        gg   = Trim(Request.form("g_id"))
                        If Not(IsNumeric(gg)) Then gg = 0
                        name = code_form(Request.form("name"))
                        url  = code_form(Request.form("url"))

                        If Len(name) < 1 Or Len(name) > 50 Or Len(url) < 8 Or Len(url) > 100 Then
                            Response.Write("<script language=javascript>" & _
                            vbcrlf & "alert(""�������ǩ�� ��ǩ���� �� ��ǩ��ַ �Ǳ���Ҫ�ģ�\n\n�뷵�������롣"");" & _
                            vbcrlf & "history.back(1)" & _
                            vbcrlf & "</script>")
                            close_conn

                            Exit Sub
                            End If

                            sql = "insert into user_bookmark(g_id,username,name,url) values(" & gg & ",'" & login_username & "','" & name & "','" & url & "')"
                            conn.execute(sql)
                            Response.Write("<script language=javascript>" & _
                            vbcrlf & "alert(""�ɹ��������һ���ҵĸ�����ǩ��" & name & """);" & _
                            vbcrlf & "</script>")
                        End Sub %>
<script language=javascript>
<!--
function group_edit(geid,gename)
{
  var gevar='������Ҫ�޸ĵ���ǩ�飨ID��'+geid+'���������ƣ����Ȳ��ܳ���20λ';
  this.document.groupedit_frm.g_id.value=geid;
  var gename=prompt(gevar+'��',gename);
  if (gename == null || gename == '' || gename.length>20)
  { alert(gevar+"��");return; }
  else
  { this.document.groupedit_frm.g_name.value=gename; }
  this.document.groupedit_frm.submit();
}

function group_del(gdid)
{
  if (confirm("�˲�����ɾ��IDΪ "+gdid+" ����ǩ�飡\n���Ҫɾ����\nɾ�����޷��ָ���"))
  { window.location="?action=groupdel&g_id="+gdid; }
}

function bookmark_edit(bid,bname,burl)
{
  var var1='������Ҫ�޸ĵĸ�����ǩ��ID��'+bid+'�������ƣ����Ȳ��ܳ���50λ';
  var var2='������Ҫ�޸ĵĸ�����ǩ��ID��'+bid+'���ĵ�ַ�����Ȳ��ܳ���100λ';
  this.document.bookmarkedit_frm.id.value=bid;
  var bename=prompt(var1+'��',bname);
  if (bename == null || bename == '' || bename.length>50)
  { alert(var1+"��");return; }
  else
  {
    this.document.bookmarkedit_frm.name.value=bename;
    var beurl=prompt(var2+'��',burl);
    if (beurl == null || beurl == '' || beurl.length>100)
    { alert(var2+"��");return; }
    else
    {this.document.bookmarkedit_frm.url.value=beurl;}
  }
  this.document.bookmarkedit_frm.submit();
}

function bookmark_del(bdid)
{
  if (confirm("�˲�����ɾ��IDΪ "+bdid+" �ĸ�����ǩ��\n���Ҫɾ����\nɾ�����޷��ָ���"))
  { window.location="?action=bookmarkdel&id="+bdid; }
}
-->
</script>
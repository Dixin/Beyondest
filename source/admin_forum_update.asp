<!-- #include file="include/onlogin.asp" -->
<!-- #include file="include/conn.asp" -->
<%
' ====================
' Beyondest.Com v3.6.1
' http://beyondest.com
' ====================

Dim admin_menu
admin_menu = "<a href='admin_forum.asp'>��̳����</a> �� " & _
"<a href='admin_forum_update.asp'>������̳����</a> �� " & _
"<a href='admin_forum.asp?action=mod'>�ϲ���̳</a> �� " & _
"<a href='admin_forum.asp?action=order'>��������</a>"
Response.Write header(11,admin_menu)

Select Case action
    Case "update_config"
        Call update_config()
    Case "update_forum"
        Call update_forum()
End Select

Sub update_config()
    Dim rs,sql,num_topic,num_data,num_reg,new_username,num_news,num_article,num_down
    num_reg = 0:num_topic = 0:num_data = 0:num_news = 0:num_article = 0:num_down = 0
    Set rs  = Server.CreateObject("adodb.recordset")
    sql     = "select username from user_data order by id desc"
    rs.open sql,conn,1,1

    If Not(rs.eof And rs.bof) Then
        num_reg      = Int(rs.recordcount)
        new_username = rs("username")
    End If

    rs.Close

    sql    = "select count(id) from bbs_topic"
    Set rs = conn.execute(sql)
    If Not(rs.eof And rs.bof) Then num_topic = Int(rs(0))
    rs.Close

    sql    = "select count(id) from bbs_data"
    Set rs = conn.execute(sql)
    If Not(rs.eof And rs.bof) Then num_data = Int(rs(0))
    rs.Close

    sql    = "select count(id) from news where hidden=1"
    Set rs = conn.execute(sql)
    If Not(rs.eof And rs.bof) Then num_news = Int(rs(0))
    rs.Close

    sql    = "select count(id) from article where hidden=1"
    Set rs = conn.execute(sql)
    If Not(rs.eof And rs.bof) Then num_article = Int(rs(0))
    rs.Close

    sql    = "select count(id) from down where hidden=1"
    Set rs = conn.execute(sql)
    If Not(rs.eof And rs.bof) Then num_down = Int(rs(0))
    rs.Close

    sql = "update configs set num_topic=" & num_topic & ",num_data=" & num_data & ",num_reg=" & num_reg & ",new_username='" & new_username & "',num_news=" & num_news & ",num_article=" & num_article & ",num_down=" & num_down & " where id=1"
    conn.execute(sql)

    Response.Write "<script language=javascript>alert(""�ɹ���������վͳ�����ݣ�"");</script>"
End Sub

Sub update_forum()
    Dim rsf,sqlf,rssum,i,rs,sql,forumid,t1,t2,t3
    sqlf        = "select * from bbs_forum order by forum_id"
    Set rsf     = conn.execute(sqlf)

    Do While Not rsf.eof
        forumid = rsf("forum_id")
        Set rs  = Server.CreateObject("adodb.recordset")
        sql     = "select * from bbs_topic where forum_id=" & forumid & " order by id desc"
        rs.open sql,conn,1,1

        If rs.eof And rs.bof Then
            t1 = 0
            t2 = "|||"
        Else
            t1 = rs.recordcount
            t2 = rs("username") & "|" & rs("tim") & "|" & rs("id") & "|" & rs("topic")
            t2 = Replace(t2,"'","")
        End If

        rs.Close:Set rs = Nothing

        sql    = "select count(*) from bbs_data where forum_id=" & forumid
        Set rs = conn.execute(sql)
        t3     = rs(0)
        rs.Close:Set rs = Nothing
        If Int(t3) < 1 Then t3 = 0

        sql = "update bbs_forum set forum_topic_num=" & t1 & ",forum_new_info='" & t2 & "',forum_data_num=" & t3 & " where forum_id=" & forumid
        conn.execute(sql)
        rsf.movenext
    Loop

    rsf.Close:Set rsf = Nothing

    Response.Write "<script language=javascript>alert(""�ɹ������˷���̳���ݣ�"");</script>"
End Sub %>
<table border=1 cellspacing=0 cellpadding=2 width=500 bordercolorlight=#C0C0C0 bordercolordark=#FFFFFF>
<tr height=50 align=center>
<td width='20%'><font class=red_2>ע������</font></td>
<td width='80%'>�����еĲ������ܽ��ǳ����ķ�������Դ�����Ҹ���ʱ��ܳ�������ϸȷ��ÿһ��������ִ�У�</td>
</tr>
<tr align=center height=80>
<td><font class=red_3>������̳������</font></td>
<td class=htd>�������İ�ť�����¼���������̳���������⡢�ظ����������¼����û�����Ϣ������ÿ��һ��ʱ������һ�Ρ�<br>
<input type=button value='����������վͳ������' onclick=update_config() class=red></td>
</tr>
<tr align=center height=80>
<td><font class=red_3>���·���̳����</font></td>
<td class=htd>�������İ�ť�����¼���ÿ����̳���������⡢�ظ��������������⡢�ظ���ʱ�����Ϣ������ÿ��һ��ʱ������һ�Ρ�<br>
<input type=button value='�������·���̳����' onclick=update_forum() class=red></td>
</tr>
<tr align=center>
<td></td>
<td></td>
</tr>
</table>
<script language=JavaScript>
<!--
function update_config()
{
if (confirm("�˲����� ���·���̳���ݣ�\n\n���Ҫ������\n\n���º��޷��ָ���"))
  window.location="admin_forum_update.asp?action=update_config"
}

function update_forum()
{
if (confirm("�˲����� ������վͳ�����ݣ�\n\n���Ҫ������\n\n���º��޷��ָ���"))
  window.location="admin_forum_update.asp?action=update_forum"
}
//-->
</script>
<%
close_conn
Response.Write ender() %>
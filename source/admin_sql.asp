<!-- #include file="include/onlogin.asp" -->
<!-- #include file="include/conn.asp" -->
<%
' ====================
' Beyondest.Com v3.6.1
' http://beyondest.com
' ====================

tit = "<a href='?'>ִ��SQL</a>"
Response.Write header(2,tit)

If Trim(Request.form("sql_run")) = "yes" Then
    Response.Write sql_chk()
Else
    Response.Write sql_type()
End If

close_conn
Response.Write ender()

Function sql_type()
    sql_type = "<table border=0><form action='admin_sql.asp' method=post><input type=hidden name=sql_run value='yes'><tr><td>������SQL��䣺��<font class=red>��ע��SQL�﷨���Լ��ٴ���</font>��&nbsp;<input type=checkbox name=is_ok value='yes'>&nbsp;�Ƿ�ȷ��</td></tr><tr><td height=50><input type=text name=sql_var size=60></td></tr><tr><td align=center><input type=submit value=' ִ �� '><font class=red_3>&nbsp;&nbsp;ִ��ĳ������󽫲����ٻָ���<br><br>��ִ��SQL�﷨ǰ����ȷ���Ƿ�һ��Ҫִ�У�</font></td></tr></form></table>"
End Function

Function sql_chk()
    On Error Resume Next
    Dim is_ok
    Dim sql_var
    is_ok       = Trim(Request.form("is_ok"))
    sql_var     = var_null(Trim(Request.form("sql_var")))

    If is_ok <> "yes" Or sql_var = "" Then
        sql_chk = "<font class=red_2>����û�ж�ִ�б���SQL������ȷ����û������SQL���</font><br><br>" & go_back:Exit Function
    End If

    If Err Then
        Err.Clear
        sql_chk = "<font class=red_2>���ղŵĲ�����ִ��SQL���ǰ����������Ĵ���<br><br>" & sql_var & "<br><br>�뷵�ؼ�顣</font><br><br>" & go_back:Exit Function
    End If

    Err.Clear
    conn.execute(sql_var)

    If Err Then
        Err.Clear
        sql_chk = "<font class=red_2>ϵͳ��ִ��SQL���ʱ������ϵͳ������Ĵ���<br><br>" & sql_var & "<br><br>�������������SQL����д�����ڣ��뷵�ؼ�顣</font><br><br>" & go_back:Exit Function
    End If

    sql_chk = "<font class=red_4>ϵͳ�ɹ���ִ��SQL��䣡</font><br><br><font class=red>" & sql_var & "</font>"
End Function %>
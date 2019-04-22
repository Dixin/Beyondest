<!-- #include file="include/onlogin.asp" -->
<!-- #include file="include/conn.asp" -->
<%
' ====================
' Beyondest.Com v3.6.1
' http://beyondest.com
' ====================

tit = "<a href='?'>执行SQL</a>"
Response.Write header(2,tit)

If Trim(Request.form("sql_run")) = "yes" Then
    Response.Write sql_chk()
Else
    Response.Write sql_type()
End If

close_conn
Response.Write ender()

Function sql_type()
    sql_type = "<table border=0><form action='admin_sql.asp' method=post><input type=hidden name=sql_run value='yes'><tr><td>请输入SQL语句：（<font class=red>请注意SQL语法，以减少错误！</font>）&nbsp;<input type=checkbox name=is_ok value='yes'>&nbsp;是否确定</td></tr><tr><td height=50><input type=text name=sql_var size=60></td></tr><tr><td align=center><input type=submit value=' 执 行 '><font class=red_3>&nbsp;&nbsp;执行某项操作后将不能再恢复！<br><br>在执行SQL语法前请先确定是否一定要执行！</font></td></tr></form></table>"
End Function

Function sql_chk()
    On Error Resume Next
    Dim is_ok
    Dim sql_var
    is_ok       = Trim(Request.form("is_ok"))
    sql_var     = var_null(Trim(Request.form("sql_var")))

    If is_ok <> "yes" Or sql_var = "" Then
        sql_chk = "<font class=red_2>您还没有对执行本次SQL语句进行确定或没有输入SQL语句</font><br><br>" & go_back:Exit Function
    End If

    If Err Then
        Err.Clear
        sql_chk = "<font class=red_2>您刚才的操作在执行SQL语句前出现了意外的错误！<br><br>" & sql_var & "<br><br>请返回检查。</font><br><br>" & go_back:Exit Function
    End If

    Err.Clear
    conn.execute(sql_var)

    If Err Then
        Err.Clear
        sql_chk = "<font class=red_2>系统在执行SQL语句时出现了系统或意外的错误！<br><br>" & sql_var & "<br><br>可能是您输入的SQL语句有错误存在！请返回检查。</font><br><br>" & go_back:Exit Function
    End If

    sql_chk = "<font class=red_4>系统成功的执行SQL语句！</font><br><br><font class=red>" & sql_var & "</font>"
End Function %>
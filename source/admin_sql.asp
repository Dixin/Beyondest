<!-- #include file="include/onlogin.asp" -->
<!-- #include file="include/conn.asp" -->
<%
' ====================
' Beyondest.Com v3.6.1
' http://beyondest.com
' ====================

tit="<a href='?'>ִ��SQL</a>"
response.write header(2,tit)

if trim(request.form("sql_run"))="yes" then
  response.write sql_chk()
else
  response.write sql_type()
end if

close_conn
response.write ender()

function sql_type()
  sql_type="<table border=0><form action='admin_sql.asp' method=post><input type=hidden name=sql_run value='yes'><tr><td>������SQL��䣺��<font class=red>��ע��SQL�﷨���Լ��ٴ���</font>��&nbsp;<input type=checkbox name=is_ok value='yes'>&nbsp;�Ƿ�ȷ��</td></tr><tr><td height=50><input type=text name=sql_var size=60></td></tr><tr><td align=center><input type=submit value=' ִ �� '><font class=red_3>&nbsp;&nbsp;ִ��ĳ������󽫲����ٻָ���<br><br>��ִ��SQL�﷨ǰ����ȷ���Ƿ�һ��Ҫִ�У�</font></td></tr></form></table>"
end function

function sql_chk()
  on error resume next
  dim is_ok,sql_var
  is_ok=trim(request.form("is_ok"))
  sql_var=var_null(trim(request.form("sql_var")))
  if is_ok<>"yes" or sql_var="" then
    sql_chk="<font class=red_2>����û�ж�ִ�б���SQL������ȷ����û������SQL���</font><br><br>"&go_back:exit function
  end if
  
  if err then
    err.clear
    sql_chk="<font class=red_2>���ղŵĲ�����ִ��SQL���ǰ����������Ĵ���<br><br>"&sql_var&"<br><br>�뷵�ؼ�顣</font><br><br>"&go_back:exit function
  end if
  
  err.clear
  conn.execute(sql_var)
  if err then
    err.clear
    sql_chk="<font class=red_2>ϵͳ��ִ��SQL���ʱ������ϵͳ������Ĵ���<br><br>"&sql_var&"<br><br>�������������SQL����д�����ڣ��뷵�ؼ�顣</font><br><br>"&go_back:exit function
  end if
  
  sql_chk="<font class=red_4>ϵͳ�ɹ���ִ��SQL��䣡</font><br><br><font class=red>"&sql_var&"</font>"
end function
%>
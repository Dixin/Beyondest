<!-- #include file="include/config.asp" -->
<!-- #include file="INCLUDE/CONN_IP.ASP" -->
<%
' ====================
' Beyondest.Com v3.6.1
' http://beyondest.com
' ====================

dim ppip,ip_ok,useres_ip,useres_port,useres_address,useres_url
ppip=trim(request.querystring("ip"))
if ip_true(ppip)="no" then
  ip_ok="no"
else
  useres_ip=ip_ip(ppip)
  useres_port=ip_port(ppip)
  useres_address=useres_ip:useres_address=ip_address(useres_address)
  ip_ok="yes"
end if

call close_conn()

useres_url=trim(request.servervariables("http_referer"))
if useres_url="" or isnull(useres_url) then
  useres_url="parent.location='/main.asp'"
else
  useres_url="history.back(1)"
end if
if useres_port<>"yes" and useres_port<>"no" and isnumeric(useres_port) then
  useres_port="�� �� �ţ�" & useres_port & "\n"
else
  useres_port=""
end if

if ip_ok="yes" then
  response.write "<script language=javascript>" & _
				 vbcrlf & "alert(""��Ҫ��ѯ�� IP ��Ϣ���£�\n\nIP ��ַ��" & useres_ip & "\n" & useres_port & "��Դ������" & useres_address & """);" & _
				 vbcrlf & useres_url & _
				 vbcrlf & "</script>"
else
  response.write "<script language=javascript>" & _
				 vbcrlf & "alert(""���Ŀ��ܽ����˷Ƿ�������\n\nϵͳ���Զ�������̳��ҳ��"");" & _
				 vbcrlf & "parent.location='main.asp'" & _
				 vbcrlf & "</script>"
end if
%>
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
  useres_port="端 口 号：" & useres_port & "\n"
else
  useres_port=""
end if

if ip_ok="yes" then
  response.write "<script language=javascript>" & _
				 vbcrlf & "alert(""您要查询的 IP 信息如下：\n\nIP 地址：" & useres_ip & "\n" & useres_port & "来源地区：" & useres_address & """);" & _
				 vbcrlf & useres_url & _
				 vbcrlf & "</script>"
else
  response.write "<script language=javascript>" & _
				 vbcrlf & "alert(""您的可能进行了非法操作！\n\n系统将自动返回论坛首页。"");" & _
				 vbcrlf & "parent.location='main.asp'" & _
				 vbcrlf & "</script>"
end if
%>
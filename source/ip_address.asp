<!-- #include file="include/config.asp" -->
<!-- #include file="include/CONN_IP.ASP" -->
<%
' ====================
' Beyondest.Com v3.6.1
' http://beyondest.com
' ====================

Dim ppip
Dim ip_ok
Dim useres_ip
Dim useres_port
Dim useres_address
Dim useres_url
ppip               = Trim(Request.querystring("ip"))

If ip_true(ppip) = "no" Then
    ip_ok          = "no"
Else
    useres_ip      = ip_ip(ppip)
    useres_port    = ip_port(ppip)
    useres_address = useres_ip:useres_address = ip_address(useres_address)
    ip_ok          = "yes"
End If

Call close_conn()

useres_url     = Trim(Request.servervariables("http_referer"))

If useres_url = "" Or IsNull(useres_url) Then
    useres_url = "parent.location='/main.asp'"
Else
    useres_url = "history.back(1)"
End If

If useres_port <> "yes" And useres_port <> "no" And IsNumeric(useres_port) Then
    useres_port = "�� �� �ţ�" & useres_port & "\n"
Else
    useres_port = ""
End If

If ip_ok = "yes" Then
    Response.Write "<script language=javascript>" & _
    vbcrlf & "alert(""��Ҫ��ѯ�� IP ��Ϣ���£�\n\nIP ��ַ��" & useres_ip & "\n" & useres_port & "��Դ������" & useres_address & """);" & _
    vbcrlf & useres_url & _
    vbcrlf & "</script>"
Else
    Response.Write "<script language=javascript>" & _
    vbcrlf & "alert(""���Ŀ��ܽ����˷Ƿ�������\n\nϵͳ���Զ�������̳��ҳ��"");" & _
    vbcrlf & "parent.location='main.asp'" & _
    vbcrlf & "</script>"
End If %>
<!-- #INCLUDE file="include/onlogin.asp" -->
<html>
<head>
<title><% Response.Write web_var(web_config,1) %> - ��̨����ϵͳ</title>
<meta http-equiv=Content-Type content=text/html; charset=gb2312>
<link rel=stylesheet href='include/beyondest.css' type=text/css>
<%
' ====================
' Beyondest.Com v3.6.1
' http://beyondest.com
' ====================

Select Case action
    Case "left"
        Call admin_left()
    Case "main"
        Call admin_main()
    Case Else
        Call admin_frame()
End Select

Sub admin_main() %><body topmargin=0 leftmargin=0 bgcolor=<% Response.Write color1 %>>
<table border=0 height='100%' width=600 align=center>
<tr height='100%' align=center><td width='30%'>
<%

    If Trim(Request.querystring("error")) = "popedom" Then
        Response.Write "<font class=red_2>������û����صĺ�̨����Ȩ��</font>"
    End If %><br><br><br>
<img src='IMAGES/SMALL/XX.GIF' border=0><br><br><br>
<font class=red>��ӭ����Ա��<font class=blue><b><% Response.Write Session("beyondest_online_admines") %></b></font>���ĵ�½</font>
</td><td width='70%'>
  <table border=1 width='100%' cellspacing=0 cellpadding=1<% Response.Write table1 %>>
  <tr><td colspan=2 align=center bgcolor=#ffffff class=red_3>���������йز���</td></tr>
  <tr><td>&nbsp;����������</td><td>&nbsp;<% Response.Write Request.ServerVariables("SERVER_NAME") %></td></tr>
  <tr><td>&nbsp;������IP��</td><td>&nbsp;<% Response.Write Request.ServerVariables("LOCAL_ADDR") %></td></tr>
  <tr><td>&nbsp;�������˿ڣ�</td><td>&nbsp;<% Response.Write Request.ServerVariables("SERVER_PORT") %></td></tr>
  <tr><td>&nbsp;������ʱ�䣺</td><td>&nbsp;<% Response.Write Now %></td></tr>
  <tr><td>&nbsp;IIS�汾��</td><td>&nbsp;<% Response.Write Request.ServerVariables("SERVER_SOFTWARE") %></td></tr>
  <tr><td>&nbsp;����������ϵͳ��</td><td>&nbsp;<% Response.Write Request.ServerVariables("OS") %></td></tr>
  <tr><td>&nbsp;�ű���ʱʱ�䣺</td><td>&nbsp;<% Response.Write Server.ScriptTimeout %> ��</td></tr>
  <tr><td>&nbsp;վ������·����</td><td>&nbsp;<% Response.Write Request.ServerVariables("APPL_PHYSICAL_PATH") %></td></tr>
  <tr><td>&nbsp;������CPU������</td><td>&nbsp;<% Response.Write Request.ServerVariables("NUMBER_OF_PROCESSORS") %> ��</td></tr>
  <tr><td>&nbsp;�������������棺</td><td>&nbsp;<% Response.Write ScriptEngine & "/" & ScriptEngineMajorVersion & "." & ScriptEngineMinorVersion & "." & ScriptEngineBuildVersion %></td></tr>
  <tr><td colspan=2 align=center bgcolor=#ffffff class=red_3>���֧���йز���</td></tr>
  <tr><td>&nbsp;���ݿ�(ADO)֧�֣�</td><td>&nbsp;<% If object_install("adodb.connection") = False Then %><font class=red><b>��</b></font> ����֧�֣�<% Else %><b>��</b> ��֧�֣�<% End If %></td></tr>
  <tr><td>&nbsp;FSO�ı���д��</td><td>&nbsp;<% If object_install("scripting.filesystemobject") = False Then %><font class=red><b>��</b></font> ����֧�֣�<% Else %><b>��</b> ��֧�֣�<% End If %></td></tr>
  <tr><td>&nbsp;Stream�ļ�����</td><td>&nbsp;<% If object_install("Adodb.Stream") = False Then %><font class=red><b>��</b></font> ����֧�֣�<% Else %><b>��</b> ��֧�֣�<% End If %></td></tr>
  <tr><td>&nbsp;Jmail���֧�֣�</td><td>&nbsp;<% If object_install("JMail.SMTPMail") = False Then %><font class=red><b>��</b></font> ����֧�֣�<% Else %><b>��</b> ��֧�֣�<% End If %></td></tr>
  <tr><td>&nbsp;CDONTS���֧�֣�</td><td>&nbsp;<% If object_install("CDONTS.NewMail") = False Then %><font class=red><b>��</b></font> ����֧�֣�<% Else %><b>��</b> ��֧�֣�<% End If %></td></tr>
  </table>
</td></tr>
</table><%
End Sub

Function object_install(strclassstring)
    On Error Resume Next
    object_install = False
    Dim xtestobj
    Err = 0
    Set xtestobj = Server.CreateObject(strclassstring)
    If Err = 0 Then object_install = True
    Set xtestobj = Nothing
    Err = 0
End Function

Sub admin_left() %><base target=main>
</head>
<script language=javascript>
<!--
function left_menu(lm)
{
  if (lm==1)
  {
    if (document.all.left_sys.style.display=='none')
    { document.all.left_sys.style.display=''; document.all.left_bm.style.display='none' }
    else
    { document.all.left_bm.style.display=''; document.all.left_sys.style.display='none' }
  }
  else
  {
    if (document.all.left_bm.style.display=='none')
    { document.all.left_bm.style.display=''; document.all.left_sys.style.display='none' }
    else
    { document.all.left_sys.style.display=''; document.all.left_bm.style.display='none' }
  }
  
}
-->
</script>
<body topmargin=0 leftmargin=0 bgcolor=<% Response.Write color1 %>><center>
<table border=0 width='100%' height='100%' cellspacing=0 cellpadding=0>
<tr><td width=155 align=center>
  <table border=0 width='100%' cellspacing=0 cellpadding=2>
  <tr><td align=center><a href='main.asp' target=_blank><% Response.Write web_var(web_config,1) %></a></td></tr>
  <tr><td align=center height=30><font class=red><b>�� ̨ �� �� ϵ ͳ</b></font></td></tr>
  <tr><td height=1 bgcolor=<% Response.Write color2 %>></td></tr>
  <tr><td align=center bgcolor=<% Response.Write web_var(web_color,2) %> style='CURSOR: HAND;' height=20 valign=bottom onclick="left_menu(1)"><font class=end><b>�� ϵ ͳ �� �� ��</b></font></td></tr>
  <tr><td height=1 bgcolor=<% Response.Write color2 %>></td></tr>
  <tr id=left_sys><td align=center>
    <table border=0 cellspacing=0 cellpadding=2>
    <tr><td><% Response.Write img_small("jt1") %><a href='admin_user.asp'>�û�����</a></td></tr>
    <tr><td><% Response.Write img_small("jt1") %><a href='admin_links.asp'>��������</a></td></tr>
    <tr><td><% Response.Write img_small("jt1") %><a href='admin_vote.asp'>�������</a></td></tr>
    <tr><td><% Response.Write img_small("jt1") %><a href='admin_nsort.asp'>�������</a></td></tr>
    <tr><td><% Response.Write img_small("jt1") %><a href='admin_update.asp'>��վ����</a></td></tr>
    <tr><td><% Response.Write img_small("jt1") %><a href='admin_data.asp'>���ݸ���</a></td></tr>
    <tr><td><% Response.Write img_small("jt1") %><a href='admin_review.asp'>���۹���</a></td></tr>
    <tr><td><% Response.Write img_small("jt1") %><a href='admin_upload_list.asp'>�ϴ�����</a></td></tr>
    <tr><td><% Response.Write img_small("jt1") %><a href='admin_config_edit.asp'>�����޸�</a></td></tr>
    <tr><td><% Response.Write img_small("jt1") %><a href='admin_sql.asp'>ִ��SQL</a></td></tr>
    <tr><td><% Response.Write img_small("jt1") %><a href='admin_popedom.asp'>Ȩ�޹���</a></td></tr>
    </table>
  </td></tr>
  <tr><td height=5></td></tr>
  <tr><td height=1 bgcolor=<% Response.Write color2 %>></td></tr>
  <tr><td align=center bgcolor=<% Response.Write web_var(web_color,2) %> style='CURSOR: HAND;' height=20 valign=bottom onclick="left_menu(2)" href="javsscript:;"><font class=end><b>�� �� �� �� �� ��</b></font></td></tr>
  <tr><td height=1 bgcolor=<% Response.Write color2 %>></td></tr>
  <tr id=left_bm style="display:none;"><td align=center>
    <table border=0 cellspacing=0 cellpadding=2>
    <tr><td><% Response.Write img_small("jt1") %><a href='admin_forum.asp'>��̳����</a></td></tr>
    <tr><td><% Response.Write img_small("jt1") %><a href='admin_news.asp'>���Ź���</a></td></tr>
    <tr><td><% Response.Write img_small("jt1") %><a href='admin_article.asp'>���¹���</a></td></tr>
    <tr><td><% Response.Write img_small("jt1") %><a href='admin_down.asp'>���ع���</a></td></tr>
    <tr><td><% Response.Write img_small("jt1") %><a href='admin_website.asp'>��վ�Ƽ�</a></td></tr>
    <tr><td><% Response.Write img_small("jt1") %><a href='admin_vouch.asp'>��Ƶ����</a></td></tr>
    <tr><td><% Response.Write img_small("jt1") %><a href='admin_flash.asp'>Flash����</a></td></tr>
    <tr><td><% Response.Write img_small("jt1") %><a href='admin_desktop.asp'>��ֽ����</a></td></tr>
    <tr><td><% Response.Write img_small("jt1") %><a href='admin_photo.asp'>������</a></td></tr>
    <tr><td><% Response.Write img_small("jt1") %><a href='gbook.asp' target=_blank>���Թ���</a></td></tr>
    </table>
  </td></tr>
  <tr><td height=5></td></tr>
  <tr><td height=1 bgcolor=<% Response.Write color2 %>></td></tr>
  <tr><td align=center height=20 valign=bottom bgcolor=<% Response.Write web_var(web_color,2) %>><a class=end href='admin_login.asp?action=logout' target=_top><b>�� �� �� �� �� ��</b></a></td></tr>
  <tr><td height=1 bgcolor=<% Response.Write color2 %>></td></tr>
  <tr><td align=center height=25><% Response.Write web_edition %></td></tr>
  <tr><td align=center><% Response.Write web_label %></td></tr>
  <tr><td height=20></td></tr>
  </table>
</td><td width=1 bgcolor=<% Response.Write web_var(web_color,3) %>></td></tr>
</table>
</body><%
End Sub

Sub admin_frame() %></head>
<frameset framespacing="0" cols="157,*" border="0" frameborder="0">
  <noframes>
<body topmargin="0" leftmargin="0">
  <p>����ҳʹ���˿�ܣ��������������֧�ֿ�ܡ������ <a href="<% Response.Write web_var(web_config,2) %>"><% Response.Write web_var(web_config,1) %></a></p>
</body>
  </noframes>
  <frame name="contents" target="main" src="admin.asp?action=left" scrolling="no" noresize>
  <frame name="main" src="admin.asp?action=main" scrolling="auto">
</frameset><%
End Sub %>
</html>
<!-- #include file="INCLUDE/config_user.asp" -->
<!-- #include file="include/jk_md5.asp" -->
<!-- #include file="include/conn.asp" -->
<%
' ====================
' Beyondest.Com v3.6.1
' http://beyondest.com
' ====================

dim err_head
tit="�޸�����"
err_head=img_small("jt0")

call web_head(2,0,0,0,0)

if int(popedom_format(login_popedom,41)) then call close_conn():call cookies_type("locked")
'------------------------------------left----------------------------------
call left_user()
'----------------------------------left end--------------------------------
call web_center(0)
'-----------------------------------center---------------------------------
response.write ukong&table1&vbcrlf&"<tr"&table2&" height=25><td height=20 class=end background=images/"&web_var(web_config,5)&"/bar_3_bg.gif>&nbsp;"&img_small(us)&"&nbsp;&nbsp;<b>�޸��ҵĸ�������</b></td></tr><tr"&table3&"><td height=150 align=center>"

sql="select * from user_data where username='"&login_username&"'"
set rs=server.createobject("adodb.recordset")
rs.open sql,conn,1,3
if rs.eof and rs.bof then
  rs.close:set rs=nothing
  call close_conn()
  call format_redirect("login.asp")
  response.end
end if

select case trim(request.form("user_edit"))
case "yes"
  response.write edit_chk()
case else
  response.write edit_main()
end select

rs.close

response.write vbcrlf&"<tr"&table2&" height=25><td height=20 class=end background=images/"&web_var(web_config,5)&"/bar_3_bg.gif><a name='pass'></a>&nbsp;"&img_small(us)&"&nbsp;&nbsp;<b>�޸��ҵĵ�½����</b></td></tr><tr"&table3&"><td height=150 align=center>"

select case trim(request("user_pass"))
case "yes"
  response.write pass_chk()
case else
  response.write pass_main()
end select

response.write vbcrlf&"</td></tr></table><br>"
'---------------------------------center end-------------------------------
call web_end(0)

function edit_main()
  dim seboy,segirl,rsface,rfs,fff:fff=0
  edit_main=edit_main & vbcrlf & "<table border=0 width='98%'>" & _
	    vbcrlf & "<form name=user_edit_frm action='?' method=post><input type=hidden name=user_edit value='yes'>" & _
	    vbcrlf & "<tr><td width='100%' colspan=3 align=center height=30><font class=red><b>ע�⣺</b></font>�����Ǻţ�" & redx & "���������Ŀ������д.</td></tr>" & _
	    vbcrlf & "<tr><td width='16%'>����ͷ�Σ�</td><td width='84%' colspan=2><input type=text name=nname value='" & code_form(rs("nname")) & "' size=28 maxlength=20></td></tr>"
  if rs("sex")=false then
    segirl=" checked":seboy=""
  else
    seboy=" checked":segirl=""
  end if
  edit_main=edit_main & vbcrlf & "<script language=javascript>function showimage(){ document.images.face_img.src=""images/face/""+document.user_edit_frm.face.options[document.user_edit_frm.face.selectedIndex].value+"".gif""; }</script>" & _
	    vbcrlf & "<tr><td width='16%'>����ձ�</td><td width='45%'> <input type=radio value=true name=sex" & seboy & " class=bg_1>&nbsp;Boy��&nbsp;<input type=radio name=sex value=false" & segirl & " class=bg_1>&nbsp;Girl</td>" & _
	    vbcrlf & "<td width='39%' align=center><a href='user_face.asp' target=_blank>���鿴����ͷ��</a>&nbsp;&nbsp;" & _
	    vbcrlf & "<select size=1 name=face style='width: 50;' onChange=""showimage()"">"
  rsface=rs("face")
  for i=0 to web_var(web_num,11)
    rfs=""
    if int(rsface)=i then rfs=" selected":fff=1
    edit_main=edit_main & vbcrlf & "<option value='" & i & "'" & rfs & ">" & i & "</option>"
  next
  if fff=0 then edit_main=edit_main & vbcrlf & "<option value='" & rsface & "' selected>" & rsface & "</option>"
  edit_main=edit_main & vbcrlf & "</select></td></tr>" & _
	    vbcrlf & "<tr><td height=30>������գ�</td><td><select name=b_year>"
  dim bires,birse:bires=rs("birthday")
  if not(isdate(bires)) then bires=#1982/6/16#
  for i=1900 to year(now)
    birse=""
    if int(year(bires))=int(i) then birse=" selected"
    edit_main=edit_main & vbcrlf & "<option value='" & i & "'" & birse & ">" & i & "</option>"
  next
  edit_main=edit_main & vbcrlf & "</select>�� <select name=b_month>"
  for i=1 to 12
    birse=""
    if int(month(bires))=int(i) then birse=" selected"
    edit_main=edit_main & vbcrlf & "<option value='" & i & "'" & birse & ">" & i & "</option>"
  next
  edit_main=edit_main & vbcrlf & "</select>�� <select name=b_day>"
  for i=1 to 31
    birse=""
    if int(day(bires))=int(i) then birse=" selected"
    edit_main=edit_main & vbcrlf & "<option value='" & i & "'" & birse & ">" & i & "</option>"
  next
  edit_main=edit_main & vbcrlf & "</select>��</td><td rowspan=5 align=center><img border=0 src='images/face/" & rsface & ".gif' name=face_img></td></tr>" & _
	    vbcrlf & "<tr><td width='16%'>�����ʼ���</td><td width='45%'><input type=text name=email value='" & rs("email") & "' size=28 maxlength=50>"&redx&"</td></tr>" & _
	    vbcrlf & "<tr><td>���QQ��</td><td><input type=text name=qq value='" & rs("qq") & "' size=28 maxlength=15></td></tr>" & _
	    vbcrlf & "<tr><td>�����ҳ��</td><td><input type=text name=url value='" & code_form(rs("url")) & "' size=28 maxlength=100></td></tr>" & _
	    vbcrlf & "<tr><td>�������</td><td><input type=text name=whe value='" & code_form(rs("whe")) & "' size=28 maxlength=20></td></tr>" & _
	    vbcrlf & "<tr><td valign=top><br>���˽��ܣ�</td><td colspan=2 valign=top>" & _ 
	    vbcrlf & "<table border=0 width='100%' cellspacing=0 cellpadding=0><tr><td width='69%'>" & _
	    vbcrlf & "<textarea rows=7 name=remark cols=42>" & rs("remark") & "</textarea></td>" & _
	    vbcrlf & "<td width='31%' valign=top><br>" & redx & "ע�⣺<br><br><br>" & web_var(web_error,3) & _       
	    vbcrlf & "</td></tr></table>" & _
	    vbcrlf & "</td></tr>" & _
	    vbcrlf & "<tr><td></td><td colspan=2 height=50>" & _
	    vbcrlf & "<input type=submit value=' �� �� �� �� '>������<input type=reset value=' �� �� �� �� '>" & _
	    vbcrlf & "</td></form></tr></table><br>"
end function

function edit_chk()
  dim nname,sex,birthday,face,email,qq,url,whe,remark,founderr
  nname=code_form(trim(request.form("nname")))
  sex=trim(request.form("sex"))
  birthday=trim(request.form("b_year"))&"-"&trim(request.form("b_month"))&"-"&trim(request.form("b_day"))
  face=trim(request.form("face"))
  email=code_form(trim(request.form("email")))
  qq=trim(request.form("qq"))
  url=code_form(trim(request.form("url")))
  whe=code_form(trim(request.form("whe")))
  remark=code_form(request.form("remark"))
  
  founderr=""
  if not(isdate(birthday)) then
    founderr=founderr&err_head&"��ѡ��� <font class=red_3>����</font> ����һ����Ч�����ڸ�ʽ��<br>"
  end if
  if email_ok(email)<>"yes" or len(email)>50 then
    founderr=founderr&err_head&"������� <font class=red_3>E-mail</font> Ϊ�ջ򲻷����ʼ�����<br>"
  end if
  if qq<>"" and not isnull(qq) then
    if not(isnumeric(qq)) or len(qq)>15 then
      founderr=founderr&err_head&"������� <font class=red_3>QQ</font> �������ֻ򳤶ȳ���15λ��<br>"
    end if
  end if
  if len(remark)>250 then
    founderr=founderr&err_head&"������� <font class=red_3>���˽���</font> ̫���ˣ����ܳ���250���ַ���<br>"
  end if
  
  if founderr="" then
    rs("nname")=nname
    rs("sex")=sex
    rs("birthday")=birthday
    rs("face")=face
    rs("email")=email
    if qq<>"" and not isnull(qq) then
      rs("qq")=qq
    end if
    rs("url")=url
    rs("whe")=whe
    rs("remark")=remark
    rs.update
    
    edit_chk="<font class=red>���ѳɹ��޸������Ļ������ϣ�</font>" & VbCrLf & "<br><br><a href='user_main.asp'>����"&tit_fir&"</a>" & vbcrLf & "<br><br>��ϵͳ���� " & web_var(web_num,5) & " ���Ӻ��Զ����أ�" & _
	     VbCrLf & "<meta http-equiv='refresh' content='" & web_var(web_num,5) & "; url=user_main.asp'>"    
    exit function
  else
    edit_chk=found_error(founderr,300):exit function
  end if
end function

function pass_main()
%>
<table border=0 width=300 cellspacing=0 cellpadding=2>
<form action='#pass' method=post>
<input type=hidden name=user_pass value='yes'>
<tr height=10><td colspan=2></td></tr>
<tr align=center>
<td width='30%'>��½���룺</td>
<td width='70%'><input type=password name=password size=25 maxlength=20></td>
</tr>
<tr align=center>
<td>�ظ����룺</td>
<td><input type=password name=password2 size=25 maxlength=20></td>
</tr>
<tr align=center>
<td>����Կ�ף�</td>
<td><input type=text name=passwd size=25 maxlength=20></td>
</tr>
<tr height=30><td colspan=2 align=center><input type=submit value=' �� �� �� �� '></td></tr>
</form>
</table>
<%
end function

function pass_chk()
  dim password,password2,passwd,founderr,rs,sql
  password=trim(request.form("password"))
  password2=trim(request.form("password2"))
  passwd=trim(request.form("passwd"))
  
  founderr=""
  if symbol_ok(password)<>"yes" then
    founderr=founderr&err_head&"������� <font class=red_3>��½����</font> Ϊ�ջ򲻷�����ع���<br>"
  else
    if password<>password2 then
      founderr=founderr&err_head&"������� <font class=red_3>��½����</font> �� <font class=founderr>ȷ������</font> ��һ�£�<br>"
    end if
  end if
  if symbol_name(passwd)<>"yes" then
    founderr=founderr&err_head&"������� <font class=red_3>����Կ��</font> Ϊ�ջ򲻷�����ع���<br>"
  end if
  
  if founderr="" then
    set rs=server.createobject("adodb.recordset")
    sql="select password,passwd from user_data where username='"&login_username&"' and password='"&login_password&"'"
    rs.open sql,conn,1,3
    if rs.eof and rs.bof then
      pass_chk="<font class=red_2>���޸Ĺ����г������˵�½��Ϣ��������⣡</font><br><br>����� <a href='help.asp?action=register' class=red_3>��Աע��ע������</a> �鿴�й�����<br><br>�� <a href='login.asp?action=logout'>����˵��?/a> ���ٴν����޸�<br><br>��ϵͳ���� " & web_var(web_num,5) & " ���Ӻ��Զ��ص�½��" & _
	       VbCrLf & "<meta http-equiv='refresh' content='" & web_var(web_num,5) & "; url=logout.asp'>"
      rs.close:set rs=nothing:exit function
    else
      password=jk_md5(password,"short")
      rs("password")=password
      rs("passwd")=jk_md5(passwd,"short")
      rs.update
      response.cookies("beyondest_online")("login_password")=password
      if request.cookies("beyondest_online")("iscookies")="yes" then
        response.cookies("beyondest_online").expires=date+365
      end if
      pass_chk="<font class=red>���ѳɹ��޸������� ��½���� �� ����Կ�ף�</font>" & VbCrLf & "<br><br><a href='user_main.asp'>�����û�����</a>" & VbCrLf & "<br><br>��ϵͳ���� " & web_var(web_num,5) & " ���Ӻ��Զ����أ�" & _
	       VbCrLf & "<meta http-equiv='refresh' content='" & web_var(web_num,5) & "; url=user_main.asp'>"
      rs.close:set rs=nothing:exit function
    end if
    rs.close:set rs=nothing
  else
    founderr=founderr&err_head&"������й� <a href='help.asp?action=register' class=red_3>��Աע��ע������</a> ��������д��"
    pass_chk=found_error(founderr,280):exit function
  end if
end function
%>
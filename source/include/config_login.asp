<!-- #include file="config.asp" -->
<!-- #include file="skin.asp" -->
<!-- #include file="jk_md5.asp" -->
<%
'*******************************************************************
'
'                     Beyondest.Com V3.6 Demo版
' 
'           网址：http://www.beyondest.com
' 
'*******************************************************************

dim err_head
err_head=img_small("jt0")
index_url="user_main"
tit_fir=format_menu(index_url)

sub nopass()
  dim pass_action
  pass_action=trim(request.form("pass_action"))
  select case pass_action
  case "question"
    if post_chk()<>"yes" then
      call close_conn()
      call cookies_type("post")
    end if
    response.write pass_question()
  case "chk"
    if post_chk()<>"yes" then
      call close_conn()
      call cookies_type("post")
    end if
    response.write pass_chk()
  case else
    response.write pass_type()
  end select
end sub

function pass_question()
  dim username
  username=trim(request.form("username"))
  if symbol_name(username)<>"yes" then
    pass_question="您输入的 <font class=red>登陆名称</font> 为空或不符合相关规则！<br><br>"&go_back
    exit function
  end if
  pass_question="<table border=0 class=fr><form action='login.asp?action=nopass' method=post><input type=hidden name=pass_action value='chk'><tr height=40><td>登陆名称：</td><td><input type=text name=uname size=20 value='"&username&"' readonly class=black_bg></td></tr><tr height=25><td>密码钥匙：</td><td><input type=password name=passwd size=20 maxlength=20></td></tr><tr height=25><td>新的密码：</td><td><input type=password name=password size=20 maxlength=20></td></tr><tr height=25><td>重复密码：</td><td><input type=password name=password2 size=20 maxlength=20></td></tr><tr height=40><td></td><td><input type=submit value='下 一 步'></td></tr><input type=hidden name=username value='"&username&"'></form></table>"
end function

function pass_chk()
  dim username,uname,passwd,password,password2
  username=trim(request.form("username"))
  uname=trim(request.form("uname"))
  passwd=trim(request.form("passwd"))
  password=trim(request.form("password"))
  password2=trim(request.form("password2"))
  if symbol_name(username)<>"yes" or username<>uname then
    pass_chk="您输入的 <font class=red>登陆名称</font> 为空或不符合相关规则！<br><br>"&go_back
    exit function
  end if
  if symbol_name(passwd)<>"yes" then
    pass_chk="您输入的 <font class=red>密码钥匙</font> 为空或不符合相关规则！<br><br>"&go_back
    exit function
  end if
  if symbol_ok(password)<>"yes" then
    pass_chk="您输入的 <font class=red>登陆密码</font> 为空或不符合相关规则！<br><br>"&go_back
    exit function
  else
    if password<>password2 then
      pass_chk="<font class=red>登陆密码</font> 和 <font class=red>确认密码</font> 不一致！<br><br>"&go_back
      exit function
    end if
  end if
  sql="select top 1 password from user_data where username='"&username&"' and passwd='"&jk_md5(passwd,"short")&"' and hidden=1"
  set rs=server.createobject("adodb.recordset")
  rs.open sql,conn,1,3
  if rs.eof and rs.bof then
    rs.close:set rs=nothing
    pass_chk="<font class=red>登陆名称</font> 和 <font class=red>密码钥匙</font> 有错或您已被锁定！<br><br>"&go_back
    exit function
  end if
  rs("password")=jk_md5(password,"short")
  rs.update
  rs.close:set rs=nothing
  pass_chk="<font class=blue_1><b>"&username&"</b></font>，<font class=red>您已成功修改了您的密码！</font><br><br>新密码是：<font class=red_3>"&password2&"</font> 请劳记！<br><br><a href='login.asp'>点击进入登陆页面</a>"
end function

function pass_type()
  pass_type="<table border=0><form action='login.asp?action=nopass' method=post><input type=hidden name=pass_action value='question'><tr height=40><td>您的登陆名称：</td><td><input type=text name=username size=20 maxlength=20></td></tr><tr height=40><td></td><td><input type=submit value='下 一 步'></td></tr></form></table>"
end function

sub register_main()
  dim reg_action,left_i
  reg_action=trim(request.form("reg_action"))
  select case reg_action
  case "reg_main"
    left_i=2
  case "reg_chk"
    left_i=3
  case else
    left_i=1
  end select
%>
<table border=0 width='100%' cellspacing=0 cellpadding=0>
<tr valign=top align=center><td width='23%'>
<br><br><br><img name=reg_left src='images/<%response.write web_var(web_config,5)%>/reg_left_<%=left_i%>.gif' border=0>
</td><td width='77%'>
  <table border=0 width='90%' cellspacing=0 cellpadding=0>
  <tr><td align=center height=80><img src='images/<%response.write web_var(web_config,5)%>/reg_top.gif' border=0></td></tr>
  <tr><td align=center height=300><%
  select case reg_action
  case "reg_main"
    call reg_type()
  case "reg_chk"
    response.write reg_chk()
  case else
    call reg_policy()
  end select
%><br><br></td></tr>
  </table>
</td></tr></table>
<%
end sub

sub reg_policy()
%>
<table border=0 width=450 cellspacing=0 cellpadding=0>
<tr><td class=htd>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;欢迎您加入本站点参加交流和讨论，本站点将向技术型站点转变。<br><br>
为维护网上公共秩序和社会稳定，请您自觉遵守以下条款：<br><br>
　一、不得利用本站危害国家安全、泄露国家秘密，不得侵犯国家社会集体的和公民的合法权益，不得利用本站制作、复制和传播下列信息： <br>
　　（一）煽动抗拒、破坏宪法和法律、行政法规实施的；<br>
　　（二）煽动颠覆国家政权，推翻社会主义制度的；<br>
　　（三）煽动分裂国家、破坏国家统一的；<br>
　　（四）煽动民族仇恨、民族歧视，破坏民族团结的；<br>
　　（五）捏造或者歪曲事实，散布谣言，扰乱社会秩序的；<br>
　　（六）宣扬封建迷信、淫秽、色情、赌博、暴力、凶杀、恐怖、教唆犯罪的；<br>
　　（七）公然侮辱他人或者捏造事实诽谤他人的，或者进行其他恶意攻击的；<br>
　　（八）损害国家机关信誉的；<br>
　　（九）其他违反宪法和法律行政法规的；<br>
　　（十）进行商业广告行为的。<br>
　二、互相尊重，对自己的言论和行为负责。</td></tr>
<form name=form_reg action='login.asp?action=register' method=post>
<input type=hidden name=reg_action value='reg_main'>
<tr><td align=center height=50>
<input type=submit value="我已阅读并同意以上条款">&nbsp;┋&nbsp;<input type=button value="不同意" onClick="document.location='index.asp'">
</td></tr>
</form>
</table>
<%
end sub

sub reg_type()
%><br>
  <table border=0 width=360 cellspacing=0 cellpadding=2>
  <tr><td width='35%'></td><td width='65%'></td></tr>
  <form name=reg_frm action='login.asp?action=register' method=post>
  <input type=hidden name=reg_action value='reg_chk'>
  <tr>
    <td align=center>用户名称：</td>
    <td><input type=text name=username size=20 maxlength=20><%=redx%></td>
  </tr>
  <tr>
    <td align=center>登陆密码：</td>
    <td><input type=password name=password size=20 maxlength=20><%=redx%></td>
  </tr>
  <tr>
    <td align=center>确认密码：</td>
    <td><input type=password name=password2 size=20 maxlength=20><%=redx%></td>
  </tr>
  <tr>
    <td align=center>密码钥匙：</td>
    <td><input type=text name=passwd size=20 maxlength=20><%=redx%></td>
  </tr>
  <tr>
    <td align=center>E - mail：</td>
    <td><input type=text name=email size=30 maxlength=50><%=redx%></td>
  </tr>
  <tr>
    <td align=center>您的性别：</td>
    <td>&nbsp;<input type=radio name=sex value='boy' checked class=bg_1>&nbsp;男孩&nbsp;&nbsp;&nbsp;<input type=radio name=sex value='girl' class=bg_1>&nbsp;女孩&nbsp;<%response.write redx%></td>
  </tr>
  <tr><td></td><td height=50><input type=submit value=' 现 在 注 册 '></td></tr>
</form>
  <tr><td colspan=2 height=30><hr size=1 color=#c0c0c0 width='90%'></td></tr>
  <tr><td colspan=2>
<p style='line-height: 150%'><font class=red_2>用户注册注意事项：</font><br>
&nbsp;&nbsp;&nbsp;1、用户名称注册申请成功之后将不能更改。<br>
&nbsp;&nbsp;&nbsp;2、用户名称可以是大小写英文字母（a~z、A~Z）、阿拉伯数字（0~9）、
连接字符“-”、下划线“_”和汉字组成；首字符不能为连接字符“-”或下划线“_”，长度不能超过20位。例：joe_527<br>
&nbsp;&nbsp;&nbsp;3、登陆密码只能由大小写英文字母（a~z、A~Z）、阿拉伯数字（0~9）、
连接字符“-”和下划线“_”组成。例：dw7v9j<br>
&nbsp;&nbsp;&nbsp;4、以上填写的注册信息的内空均不区分大小写。</p>
  </td></tr>
  </table>
<%
end sub

function reg_chk()
  dim username,password,password2,passwd,email,red
  username=trim(request.form("username"))
  password=trim(request.form("password"))
  password2=trim(request.form("password2"))
  passwd=trim(request.form("passwd"))
  email=code_form(trim(request.form("email")))
  red=""
  if symbol_name(username)<>"yes" then
    red=red&err_head&"您输入的 <font class=red>用户名称</font> 为空或不符合相关规则！<br>"
  else
    if health_name(username)<>"yes" then
      red=red&err_head&"您输入的 <font class=red>用户名称</font> 含有<font class=red>本系统禁用字符</font>！<br>"
    end if
  end if
  if symbol_ok(password)<>"yes" then
    red=red&err_head&"您输入的 <font class=red>登陆密码</font> 为空或不符合相关规则！<br>"
  else
    if password<>password2 then
      red=red&err_head&"您输入的 <font class=red>登陆密码</font> 和 <font class=red>确认密码</font> 不一致！<br>"
    end if
  end if
  if symbol_name(passwd)<>"yes" then
    red=red&err_head&"您输入的 <font class=red>密码钥匙</font> 为空或不符合相关规则！<br>"
  end if
  if email_ok(email)<>"yes" or len(email)>50 then
    red=red&err_head&"您输入的 <font class=red>E-mail</font> 为空或不符合邮件规则！<br>"
  end if

  if red="" then
    sql="select * from user_data where username='" & username & "'"
    set rs=server.createobject("adodb.recordset")
    rs.open sql,conn,1,3
    if rs.eof and rs.bof then
      rs.addnew
      rs("username")=username
      rs("password")=jk_md5(password,"short")
      rs("passwd")=jk_md5(passwd,"short")
      rs("email")=email
      if trim(request.form("sex"))="girl" then
        rs("sex")=0
      else
        rs("sex")=1
      end if
      rs("face")="0"
      rs("tim")=now_time
      rs("power")="user"
      if web_var_num(web_setup,2,1)=0 then
        rs("hidden")=false
      else
        rs("hidden")=true
      end if
      rs("bbs_counter")=0
      rs("counter")=0
      rs("integral")=0
      rs("emoney")=0
      rs("login_num")=0
      rs("last_tim")=now_time
      rs("popedom")="00000000000000000000000000000000000000000000000000"
      rs.update
      rs.close:set rs=nothing
      
      conn.execute("update configs set new_username='"&username&"',num_reg=num_reg+1 where id=1")
      call reg_msg(username)
      
      if web_var_num(web_setup,2,1)=0 then
        reg_chk="<font class=red><b>"&username&"</b></font>，您已成功注册成为本站用户。<br><br>您现在的状态为：<font class=red_3>未审核</font>，请等待管理员的审核。谢谢！"
      else
        reg_chk="恭喜！<font class=red><b>"&username&"</b></font>，您已成功注册成为本站用户。<br><br><a href='login.asp'>现在进行登陆</a><br><br>请尽快登陆并修改您的个人资料。"
      end if
      exit function
    else
      red=err_head&"您输入的 <font class=red>用户名称（<b>"&username&"</b>）</font> 已经被注册！<br>" & _
	       err_head&"请重新选择输入您的 <font class=red>用户名称</font> 以并继续注册！<br>"
      rs.close:set rs=nothing
      reg_chk=found_error(red,300):exit function
    end if
    rs.close:set rs=nothing
  else
    red=red&err_head&"请查阅有关 <a href='help.asp?action=register' class=red_3>用户注册注意事项</a> 并重新填写。"
    reg_chk=found_error(red,280):exit function
  end if
end function

sub reg_msg(accept_u)
  dim msg_topic,msg_word
  msg_topic=web_var(web_config,1)&" 欢迎您的到来！"
  msg_word=web_var(web_config,1)&"全体用户和工作人员欢迎您的到来！[br]" & _
	   "如有任何疑问请及时联系我们。[br]" & _
	   "如有任何使用上的问题请查看网站帮助。[br]" & _
	   "感谢您访问本站，让我们一起来建设这个网上家园！"
  sql="insert into user_mail(send_u,accept_u,topic,word,tim,types,isread) " & _
      "values('"&web_var(web_config,3)&"','"&accept_u&"','"&msg_topic&"','"&msg_word&"','"&now_time&"',1,0)"
  conn.execute(sql)
end sub

sub login_chk()
  dim username,password,red,id,power,hidden,face
  if symbol_name(login_username)="yes" then
    username=login_username
  else
    username=trim(request.form("username"))
  end if
  if symbol_ok(login_password)="yes" then
    password=login_password
  else
    password=trim(request.form("password"))
    password=jk_md5(password,"short")
  end if

  red=""
  if symbol_name(username)<>"yes" then
    red=red&err_head&"您输入的 <font class=red_3>用户名称</font> 为空或不符合相关规则！<br>"
  end if
  if symbol_ok(password)<>"yes" then
    red=red&err_head&"您输入的 <font class=red_3>登陆密码</font> 为空或不符合相关规则！<br>"
  end if

  if red="" then
    sql="select top 1 id,face,power,hidden from user_data where username='"&username&"' and password='"&password&"'"
    set rs=conn.execute(sql)
    if rs.eof and rs.bof then
      red=err_head&"您输入的 <font class=red>用户名称</font> 和 <font class=red>登陆密码</font>  有错误！<br>" & _
	       err_head&"请重新输入以并继续登陆本站！"
      response.write found_error(red,260)
    else
      power=rs("power"):hidden=rs("hidden")
      if hidden=true then
        'response.cookies(web_cookies).path=web_path
        response.cookies(web_cookies)("login_username")=username
        response.cookies(web_cookies)("login_password")=password

        sql="update user_data set last_tim='"&now_time&"' where username='"&username&"'"
        conn.execute(sql)
        tit=request.cookies(web_cookies)("guest_name")
        if var_null(tit)<>"" then
          conn.execute("delete from user_login where l_username='"&tit&"'")
        end if
        
        if trim(request.form("memery_info"))="yes" then
	  response.cookies(web_cookies)("iscookies")="yes"
	  response.cookies(web_cookies).expires=date+365
        end if
	'----------------------------------------------------------------------------
	if trim(request.form("re_log"))="yes" then
	  call close_conn()
	  response.redirect request.servervariables("http_referer")
	  response.end
	end if
	'----------------------------------------------------------------------------
	response.write "<meta http-equiv='refresh' content='"&web_var(web_num,5)&"; url="&index_url&".asp'><br><br><br>你好，<font class=red>"&username&"</font>&nbsp;你现在是 <font class=red>"&format_power(power,1)&"</font> 登陆模式<br><br>" & _
		       vbcrlf & "<a href='"&index_url&".asp'>进入"&tit_fir&"</a>&nbsp;┋&nbsp;<a href='login.asp?action=logout'>退出本次登陆</a><br><br><br>"
      else
        response.write "<font class=red>您的用户帐号还未审核！</font><br><br>请与管理员联系。"
      end if
    end if
    rs.close
  else
    red=red&err_head&"请查阅有关 <a href='help.asp?action=register' class=red_3>用户注册注意事项</a> 并重新填写。"
    response.write found_error(red,280)
  end if
end sub

sub login_main()
%>
<script language=javascript src='style/login.js'></script>
<table border=0>
<form name=login_frm method=post action='login.asp?action=login_chk' onsubmit="return login_true()">
<tr><td align=center height=30>用户名称：&nbsp;<input type=text name=username size=15 maxlength=20 tabindex=1>&nbsp;&nbsp;</td></tr>
<tr><td align=center>登陆密码：&nbsp;<input type=password name=password size=15 maxlength=20 tabindex=2>&nbsp;&nbsp;</td></tr>
<tr><td align=center height=30 align=center><input type=radio name=memery_info value='no' class=bg_1 checked>&nbsp;登陆一次&nbsp;
<input type=radio name=memery_info value='yes' class=bg_1>&nbsp;永久登陆</td></tr>
<tr><td align=center>
<input type=button value='注 册' onClick="document.location='login.asp?action=register'">&nbsp;&nbsp;
<input type=button value='忘记密码' onClick="document.location='login.asp?action=nopass'">&nbsp;&nbsp;
<input type=submit value='登 陆' tabindex=3>
</td></tr>
</table>
<%
end sub
%>
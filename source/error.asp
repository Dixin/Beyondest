<!-- #include file="INCLUDE/config_other.asp" -->
<%
'*******************************************************************
'
'                     Beyondest.Com V3.6 Demo版
' 
'           网址：http://www.beyondest.com
' 
'*******************************************************************

dim old_url,web_skin:web_skin=web_var(web_config,5)
index_url="error"
tit="出错信息提示"
tit_fir=""

action=trim(request.cookies("beyondest_online")("error_action"))
old_url=trim(request.cookies("beyondest_online")("old_url"))
if var_null(old_url)="" then
  old_url="main.asp"
end if

call web_head(0,0,0,0,0)
'------------------------------------left----------------------------------
response.write left_action("jt13",4)
'----------------------------------left end--------------------------------
call web_center(0)
'-----------------------------------center---------------------------------
select case action
case "loading"
  tit="<font class=red>对不起，本站正在维护或更新中……<br><br>您暂时不能注册或登陆本站！<br><br>请稍等片刻……<br><br>给您带来的不便，还请见谅！！！</font>"
case "username"
  tit="<font class=red>您所查看详细的用户信息的用户名不符合有关规则或不存在！</font><br><br>请勿乱给本站的程序提交非法参数。"
case "login"
  tit="<font class=red>您可能没有注册和登陆本站或登陆信息有误！</font><br><br><font class=red_3>为支持本站的发展，正视本站成员的劳动成果！<br>本站的大部分资源（论坛、文栏、下载、短信等功能服务）<br>需要注册并正确登陆才能进行。"
case "power"
  tit="<font class=red>您的权限太低！系统不充许您进行刚才的操作！<br>可能是您要查看的软件、文章以及论坛主题等所需级别较高。</font><br><br>请勿乱给本站的程序提交非法参数。"
case "locked"
  tit="<font class=red>您的目前已被网站管理员锁定，只能进行登陆和浏览等操作！<br>原因可能是您之前进行了不友好的操作。如要解除锁定，请与网站管理员联系。</font><br><br>请勿乱给本站的程序提交非法参数。"
case "post"
  tit=post_error&"<br><br>请勿乱给本站的程序提交非法参数。"
case "effect_id"
  tit="<font class=red>您所查看的特效ID不符合有关规则或不存在！</font><br><br>请勿乱给本站的程序提交非法参数。"
case "islock"
  tit="<font class=red>您所回复的贴子已被锁定！</font><br><br>您不可以再对该贴进行回复操作。"
case "mail_id"
  tit="<font class=red>您所查看、回复、转发或删除的短信ID不符合有关规则或不存在！</font><br><br>请勿乱给本站的程序提交非法参数。"
case "edit_id"
  tit="<font class=red>您所编辑的贴子ID不符合有关规则或不存在！</font><br><br>请勿乱给本站的程序提交非法参数。"
case "del_id"
  tit="<font class=red>您所删除的贴子ID不符合有关规则或不存在！</font><br><br>请勿乱给本站的程序提交非法参数。"
case "forum_id"
  tit="<font class=red>您所查看或发表贴子的论坛ID不符合有关规则或不存在！<br>可能该贴已经被删除或该论坛已经被暂时关闭！</font><br><br>请勿乱给本站的程序提交非法参数。"
case "time_load"
  tit="<font class=red>本站已开启防刷新机制，请勿在 "&web_var(web_num,16)&" 秒钟内重复发表！</font><br><br>请勿乱给本站的程序提交非法参数。"
case "view_id"
  tit="<font class=red>您所查看或发表回贴的主题贴子ID不符合有关规则或不存在！</font><br><br>请勿乱给本站的程序提交非法参数。"
case else
  tit="<font class=red>出现未知错误！</font>请与管理员联系！<br>您尚未 <a href=login.asp?action=register>注册</a> 或者 <a href=login.asp>登陆</a>，或者不具备使用当前功能的权限。<br><br><a href='gbook.asp?action=write'>〖 告诉我们 〗</a>"
end select

if action<>"loading" then
  tit=tit&"<br><br><br><a href='"&old_url&"'>点击此处可返回出错页的前一页</a>"
end if

'response.cookies("beyondest_online")("error_action")=""
%>
<!-----------------------------------center---------------------------------->
<table border=0 width=590 cellspacing=0 cellpadding=0 class=fr>
<tr><td align=right><img src='images/<%response.write web_skin%>/center_error.gif' border=0></td></tr>
<tr><td align=center height=380>
<table border=0 cellpadding=0 cellspacing=0 width=534>
  <tr>
   <td colspan=3><img src=images/<%response.write web_skin%>/error_r1_c1.gif width=534 height=42 border=0></td>
   <td><img src=images/error/spacer.gif width=1 height=42 border=0></td>
  </tr>
  <tr>
   <td rowspan=2><img src=images/<%response.write web_skin%>/error_r2_c1.gif width=43 height=239 border=0></td>
   <td width=479 height=228 align=center bgcolor=#f7f7f7 class=htd><%response.write tit%></td>
   <td rowspan=2><img src=images/<%response.write web_skin%>/error_r2_c3.gif width=12 height=239 border=0></td>
   <td><img src=images/<%response.write web_skin%>/spacer.gif width=1 height=228 border=0></td>
  </tr>
  <tr>
   <td><img src=images/<%response.write web_skin%>/error_r3_c2.gif width=479 height=11 border=0></td>
   <td><img src=images/<%response.write web_skin%>/spacer.gif width=1 height=11 border=0></td>
  </tr>
</table><br>
</td></tr></table>
<%
'---------------------------------center end-------------------------------
call web_end(1)
%>
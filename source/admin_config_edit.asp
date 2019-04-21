<!-- #INCLUDE file="include/onlogin.asp" -->
<!-- #INCLUDE file="include/fso_file.asp" -->
<!-- #INCLUDE file="include/common_other.asp" -->
<%
' ====================
' Beyondest.Com v3.6.1
' http://beyondest.com
' ====================

Dim tit_menu
tit_menu = "<a href='?'>配置修改</a>&nbsp;┋&nbsp;" & _
"<a href='javascript:Do_b_data();'>备份配置</a>&nbsp;┋&nbsp;" & _
"<a href='javascript:Do_h_data();'>还原配置</a>"
Response.Write header(3,tit_menu) %>
<script language=JavaScript><!--
function Do_b_data()
{
if (confirm("此操作将 备份 现有的网站配置！\n\n真的要进行吗？\n备份后将无法恢复！"))
  window.location="?action=b"
}
function Do_h_data()
{
if (confirm("此操作将 还原 现有的网站配置！\n\n真的要进行吗？\n还原后将无法恢复！"))
  window.location="?action=h"
}
//--></script>
<%
If action = "b" Or action = "h" Then Call config_bh(action)

Select Case Trim(Request.querystring("edit"))
    Case "chk"
        Call config_chk()
    Case Else
        Call config_main()
End Select

Response.Write ender()

Sub config_main()
    Dim t1
    Dim t2
    Dim j
    Dim tdim
    Dim udim
    Dim tt:tt = 0 %>
<table border=1 width='92%' cellspacing=0 cellpadding=2<% Response.Write table1 %>>
<tr><td colspan=3 align=center>
  <table border=0 cellspacing=0 cellpadding=3>
  <tr>
  <td><% Response.Write img_small("jt1") %><a href='?action=config'>基本信息设置</a></td>
  <td><% Response.Write img_small("jt1") %><a href='?action=config2'>基本功能设置</a></td>
  <td><% Response.Write img_small("jt1") %><a href='?action=num'>页显示数量等</a></td>
  <td><% Response.Write img_small("jt12") %><a href='?action=info'>过滤提示设置</a></td>
  <td><% Response.Write img_small("jt1") %><a href='?action=down_up'>下载上传设置</a></td>
  </tr>
  <tr>
  <td><% Response.Write img_small("jt1") %><a href='?action=menu'>栏目菜单设置</a></td>
  <td><% Response.Write img_small("jt1") %><a href='?action=color'>网站颜色设置</a></td>
  <td><% Response.Write img_small("jt12") %><a href='?action=user'>用 户 组管理</a></td>
  <td><% Response.Write img_small("jt12") %><a href='?action=grade'>用户等级管理</a></td>
  <td><% Response.Write img_small("jt12") %><a href='?action=forum'>论坛分类管理</a></td>
  </tr>
  </table>
</td></tr>
<tr align=center bgcolor=<% Response.Write color2 %>>
<td width='16%'>设置名称</td>
<td width='50%'>参数</td>
<td width='34%'>相关说明</td>
</tr>
<form action='?action=<% Response.Write action %>&edit=chk' method=post>
<input type=hidden name=web_news_art value='<% Response.Write web_news_art %>'>
<input type=hidden name=web_shop value='<% Response.Write web_shop %>'>
<%

    If action = "config" Then
        tt = 1 %>
<tr>
<td>网站名称：</td>
<td><input type=text name=web_config_1 value='<% Response.Write web_var(web_config,1) %>' size=38 maxlength=20></td>
<td class=gray>&nbsp;</td>
</tr>
<tr>
<td>网站地址：</td>
<td><input type=text name=web_config_2 value='<% Response.Write web_var(web_config,2) %>' size=38 maxlength=50></td>
<td class=gray>&nbsp;</td>
</tr>
<tr>
<td>超级管理员：</td>
<td><input type=text name=web_config_3 value='<% Response.Write web_var(web_config,3) %>' size=20 maxlength=20></td>
<td class=gray>&nbsp;</td>
</tr>
<tr>
<td>所在目录：</td>
<td><input type=text name=web_config_4 value='<% Response.Write web_var(web_config,4) %>' size=20 maxlength=20></td>
<td class=gray>&nbsp;</td>
</tr>
<tr>
<td>网站SKIN：</td>
<td><input type=text name=web_config_5 value='<% Response.Write web_var(web_config,5) %>' size=20 maxlength=20></td>
<td class=gray>&nbsp;</td>
</tr>
<tr>
<td>Cookies名称：</td>
<td><input type=text name=web_cookies value='<% Response.Write web_cookies %>' size=20 maxlength=20></td>
<td class=gray>&nbsp;</td>
</tr>
<tr>
<td>数据库名称：</td>
<td><input type=text name=web_config_6 value='<% Response.Write web_var(web_config,6) %>' size=38 maxlength=50></td>
<td class=gray>&nbsp;</td>
</tr>
<tr>
<td>网站背景：</td>
<td><input type=text name=web_config_7 value='<% Response.Write web_var(web_config,7) %>' size=20 maxlength=20></td>
<td background='images/<% Response.Write web_var(web_config,7) %>.gif'>&nbsp;</td>
</tr>
<tr>
<td>虚拟货币：</td>
<td><input type=text name=web_config_8 value='<% Response.Write web_var(web_config,8) %>' size=20 maxlength=20></td>
<td class=gray>&nbsp;尽量减短其名称</td>
</tr>
<% Else %>
<input type=hidden name=web_config_1 value='<% Response.Write web_var(web_config,1) %>'>
<input type=hidden name=web_config_2 value='<% Response.Write web_var(web_config,2) %>'>
<input type=hidden name=web_config_3 value='<% Response.Write web_var(web_config,3) %>'>
<input type=hidden name=web_config_4 value='<% Response.Write web_var(web_config,4) %>'>
<input type=hidden name=web_config_5 value='<% Response.Write web_var(web_config,5) %>'>
<input type=hidden name=web_cookies value='<% Response.Write web_cookies %>'>
<input type=hidden name=web_config_6 value='<% Response.Write web_var(web_config,6) %>'>
<input type=hidden name=web_config_7 value='<% Response.Write web_var(web_config,7) %>'>
<input type=hidden name=web_config_8 value='<% Response.Write web_var(web_config,8) %>'>
<%
    End If

    If action = "config2" Then
        tt = 1 %>
<tr>
<td>网站状态：</td>
<td><input type=radio name=web_login value='1'<% If Int(web_login) = 1 Then Response.Write " checked" %> class=bg_1>&nbsp;开放&nbsp;<input type=radio name=web_login value='0'<% If Int(web_login) <> 1 Then Response.Write " checked" %> class=bg_1>&nbsp;关闭</td>
<td class=gray>&nbsp;是否开放网站</td>
</tr>
<% t1 = web_var_num(web_setup,1,1) %>
<tr>
<td>登陆浏览：</td>
<td colspan=2><input type=radio name=web_setup_1 value='1'<% If t1 = 1 Then Response.Write " checked" %> class=bg_1>&nbsp;是&nbsp;&nbsp;&nbsp;<input type=radio name=web_setup_1 value='0'<% If t1 <> 1 Then Response.Write " checked" %> class=bg_1>&nbsp;否&nbsp;&nbsp;&nbsp;&nbsp;<font class=gray>是否要登陆才可浏览文章或下载软件等</font></td>
</tr>
<% t1 = web_var_num(web_setup,2,1) %>
<tr>
<td>注册审核：</td>
<td><input type=radio name=web_setup_2 value='0'<% If t1 <> 1 Then Response.Write " checked" %> class=bg_1>&nbsp;是&nbsp;&nbsp;&nbsp;<input type=radio name=web_setup_2 value='1'<% If t1 = 1 Then Response.Write " checked" %> class=bg_1>&nbsp;否</td>
<td class=gray>&nbsp;是否对新注册进行审核</td>
</tr>
<%
        t1 = web_var_num(web_setup,3,1)
        If t1 <> 1 And t1 <> 2 Then t1 = 0 %>
<tr>
<td>网站模式：</td>
<td colspan=2 class=gray><input type=radio name=web_setup_3 value='0'<% If t1 = 0 Then Response.Write " checked" %> class=bg_1>&nbsp;<input type=text name=web_stamp_1 value='<% Response.Write web_var(web_stamp,1) %>' size=10 maxlength=10>&nbsp;&nbsp;&nbsp;本站的注册用户可以登陆，不记录在线列表</td>
<tr>
<td>&nbsp;</td>
<td colspan=2 class=gray><input type=radio name=web_setup_3 value='1'<% If t1 = 1 Then Response.Write " checked" %> class=bg_1>&nbsp;<input type=text name=web_stamp_2 value='<% Response.Write web_var(web_stamp,2) %>' size=10 maxlength=10>&nbsp;&nbsp;&nbsp;所有登陆和浏览本站的人被并记录在线列表</td>
<tr>
<td>&nbsp;</td>
<td colspan=2 class=gray><input type=radio name=web_setup_3 value='2'<% If t1 = 2 Then Response.Write " checked" %> class=bg_1>&nbsp;<input type=text name=web_stamp_3 value='<% Response.Write web_var(web_stamp,3) %>' size=10 maxlength=10>&nbsp;&nbsp;&nbsp;本站的注册用户可以登陆，并记录用户在线列表</td>
</tr>
<% t1 = web_var_num(web_setup,4,1) %>
<tr>
<td>信息过滤：</td>
<td><input type=radio name=web_setup_4 value='1'<% If t1 = 1 Then Response.Write " checked" %> class=bg_1>&nbsp;是&nbsp;&nbsp;&nbsp;<input type=radio name=web_setup_4 value='0'<% If t1 <> 1 Then Response.Write " checked" %> class=bg_1>&nbsp;否</td>
<td class=gray>&nbsp;是否对论坛等进行字符过滤</td>
</tr>
<%
        t1 = web_var_num(web_setup,5,1)
        If t1 <> 0 And t1 <> 1 Then t1 = 2 %>
<tr>
<td>显示 IP：</td>
<td><input type=radio name=web_setup_5 value='0'<% If t1 = 0 Then Response.Write " checked" %> class=bg_1> 完全保密
<input type=radio name=web_setup_5 value='1'<% If t1 = 1 Then Response.Write " checked" %> class=bg_1> 显示部分
<input type=radio name=web_setup_5 value='2'<% If t1 = 2 Then Response.Write " checked" %> class=bg_1> 完全开放</td>
<td class=gray>&nbsp;对管理员总是完全开放</td>
</tr>
<% t1 = web_var_num(web_setup,6,1) %>
<tr>
<td>版主显示：</td>
<td><input type=radio name=web_setup_6 value='1'<% If t1 = 1 Then Response.Write " checked" %> class=bg_1>&nbsp;文字链接&nbsp;<input type=radio name=web_setup_6 value='0'<% If t1 <> 1 Then Response.Write " checked" %> class=bg_1>&nbsp;下拉菜单框</td>
<td class=gray>&nbsp;论坛版主显示模式</td>
</tr>
<% t1 = web_var_num(web_setup,7,1) %>
<tr>
<td>计数方式：</td>
<td><input type=radio name=web_setup_7 value='1'<% If t1 = 1 Then Response.Write " checked" %> class=bg_1>&nbsp;记录多次&nbsp;<input type=radio name=web_setup_7 value='0'<% If t1 <> 1 Then Response.Write " checked" %> class=bg_1>&nbsp;记录一次</td>
<td class=gray>&nbsp;网站计数的方式</td>
</tr>
<% t1 = web_var_num(web_var(web_config,9),1,1) %><tr>
<td>发布新闻：</td>
<td><input type=radio name=web_config_9_1 value='1'<% If t1 = 1 Then Response.Write " checked" %> class=bg_1>&nbsp;开放&nbsp;<input type=radio name=web_config_9_1 value='0'<% If t1 = 0 Then Response.Write " checked" %> class=bg_1>&nbsp;关闭</td>
<td class=gray>&nbsp;</td>
</tr><% t1 = web_var_num(web_var(web_config,9),2,1) %>
<tr>
<td>发表文章：</td>
<td><input type=radio name=web_config_9_2 value='1'<% If t1 = 1 Then Response.Write " checked" %> class=bg_1>&nbsp;开放&nbsp;<input type=radio name=web_config_9_2 value='0'<% If t1 = 0 Then Response.Write " checked" %> class=bg_1>&nbsp;关闭</td>
<td class=gray>&nbsp;</td>
</tr><% t1 = web_var_num(web_var(web_config,9),3,1) %>
<tr>
<td>添加音乐：</td>
<td><input type=radio name=web_config_9_3 value='1'<% If t1 = 1 Then Response.Write " checked" %> class=bg_1>&nbsp;开放&nbsp;<input type=radio name=web_config_9_3 value='0'<% If t1 = 0 Then Response.Write " checked" %> class=bg_1>&nbsp;关闭</td>
<td class=gray>&nbsp;</td>
</tr><% t1 = web_var_num(web_var(web_config,9),4,1) %>
<tr>
<td>上传贴图：</td>
<td><input type=radio name=web_config_9_4 value='1'<% If t1 = 1 Then Response.Write " checked" %> class=bg_1>&nbsp;开放&nbsp;<input type=radio name=web_config_9_4 value='0'<% If t1 = 0 Then Response.Write " checked" %> class=bg_1>&nbsp;关闭</td>
<td class=gray>&nbsp;</td>
</tr><% t1 = web_var_num(web_var(web_config,9),5,1) %>
<tr>
<td>推荐网站：</td>
<td><input type=radio name=web_config_9_5 value='1'<% If t1 = 1 Then Response.Write " checked" %> class=bg_1>&nbsp;开放&nbsp;<input type=radio name=web_config_9_5 value='0'<% If t1 = 0 Then Response.Write " checked" %> class=bg_1>&nbsp;关闭</td>
<td class=gray>&nbsp;</td>
</tr>
<% Else %>
<input type=hidden name=web_login value='<% Response.Write web_login %>'>
<input type=hidden name=web_stamp_1 value='<% Response.Write web_var(web_stamp,1) %>'>
<input type=hidden name=web_stamp_2 value='<% Response.Write web_var(web_stamp,2) %>'>
<input type=hidden name=web_stamp_3 value='<% Response.Write web_var(web_stamp,3) %>'>
<input type=hidden name=web_setup_1 value='<% Response.Write web_var_num(web_setup,1,1) %>'>
<input type=hidden name=web_setup_2 value='<% Response.Write web_var_num(web_setup,2,1) %>'>
<input type=hidden name=web_setup_3 value='<% Response.Write web_var_num(web_setup,3,1) %>'>
<input type=hidden name=web_setup_4 value='<% Response.Write web_var_num(web_setup,4,1) %>'>
<input type=hidden name=web_setup_5 value='<% Response.Write web_var_num(web_setup,5,1) %>'>
<input type=hidden name=web_setup_6 value='<% Response.Write web_var_num(web_setup,6,1) %>'>
<input type=hidden name=web_setup_7 value='<% Response.Write web_var_num(web_setup,7,1) %>'>
<input type=hidden name=web_config_9_1 value='<% Response.Write web_var_num(web_var(web_config,9),1,1) %>'>
<input type=hidden name=web_config_9_2 value='<% Response.Write web_var_num(web_var(web_config,9),2,1) %>'>
<input type=hidden name=web_config_9_3 value='<% Response.Write web_var_num(web_var(web_config,9),3,1) %>'>
<input type=hidden name=web_config_9_4 value='<% Response.Write web_var_num(web_var(web_config,9),4,1) %>'>
<input type=hidden name=web_config_9_5 value='<% Response.Write web_var_num(web_var(web_config,9),5,1) %>'>
<%
    End If

    If action = "num" Then
        tt = 1 %>
<tr>
<td>用户名长度：</td>
<td><input type=text name=web_num_1 value='<% Response.Write web_var(web_num,1) %>' size=20 maxlength=20></td>
<td class=gray>&nbsp;</td>
</tr>
<tr>
<td>每页主题数：</td>
<td><input type=text name=web_num_2 value='<% Response.Write web_var(web_num,2) %>' size=20 maxlength=20></td>
<td class=gray>&nbsp;论坛、会员每页显示等</td>
</tr>
<tr>
<td>每页显示数：</td>
<td><input type=text name=web_num_3 value='<% Response.Write web_var(web_num,3) %>' size=20 maxlength=20></td>
<td class=gray>&nbsp;查看贴子内容等</td>
</tr>
<tr>
<td>每页留言数：</td>
<td><input type=text name=web_num_4 value='<% Response.Write web_var(web_num,4) %>' size=20 maxlength=20></td>
<td class=gray>&nbsp;每页留言数目等</td>
</tr>
<tr>
<td>自动返回：</td>
<td><input type=text name=web_num_5 value='<% Response.Write web_var(web_num,5) %>' size=20 maxlength=20></td>
<td class=gray>&nbsp;自动返回时间，单位为秒</td>
</tr>
<tr>
<td>主题长度：</td>
<td><input type=text name=web_num_6 value='<% Response.Write web_var(web_num,6) %>' size=20 maxlength=20>&nbsp;KB</td>
<td class=gray>&nbsp;论坛发贴主题等长度</td>
</tr>
<tr>
<td>图片宽度：</td>
<td><input type=text name=web_num_7 value='<% Response.Write web_var(web_num,7) %>' size=20 maxlength=20>&nbsp;像素</td>
<td class=gray>&nbsp;下载、图库等图片显示宽度</td>
</tr>
<tr>
<td>图片高度：</td>
<td><input type=text name=web_num_8 value='<% Response.Write web_var(web_num,8) %>' size=20 maxlength=20>&nbsp;像素</td>
<td class=gray>&nbsp;下载、图库等图片显示高度</td>
</tr>
<tr>
<td>最大宽度：</td>
<td><input type=text name=web_num_9 value='<% Response.Write web_var(web_num,9) %>' size=20 maxlength=20>&nbsp;像素</td>
<td class=gray>&nbsp;贴图、FLASH等最大宽度</td>
</tr>
<tr>
<td>最大高度：</td>
<td><input type=text name=web_num_10 value='<% Response.Write web_var(web_num,10) %>' size=20 maxlength=20>&nbsp;像素</td>
<td class=gray>&nbsp;贴图、FLASH等最大高度</td>
</tr>
<tr>
<td>用户头像数：</td>
<td><input type=text name=web_num_11 value='<% Response.Write web_var(web_num,11) %>' size=20 maxlength=20></td>
<td class=gray>&nbsp;用户头像的总数</td>
</tr>
<tr>
<td>间距系数：</td>
<td><input type=text name=web_num_12 value='<% Response.Write web_var(web_num,12) %>' size=20 maxlength=20></td>
<td class=gray>&nbsp;表格行的间距</td>
</tr>
<tr>
<td>登陆超时：</td>
<td><input type=text name=web_num_13 value='<% Response.Write web_var(web_num,13) %>' size=20 maxlength=20>&nbsp;分钟</td>
<td class=gray>&nbsp;用户登陆超时的时间</td>
</tr>
<tr>
<td>积分换算：</td>
<td><input type=text name=web_num_14 value='<% Response.Write web_var(web_num,14) %>' size=20 maxlength=20></td>
<td class=gray>&nbsp;积分换算比率</td>
</tr>
<tr>
<td>发布加分：</td>
<td><input type=text name=web_num_15 value='<% Response.Write web_var(web_num,15) %>' size=20 maxlength=20></td>
<td class=gray>&nbsp;前台发布信息加分值</td>
</tr>
<tr>
<td>防刷时间：</td>
<td><input type=text name=web_num_16 value='<% Response.Write web_var(web_num,16) %>' size=20 maxlength=20></td>
<td class=gray>&nbsp;单位为：秒</td>
</tr>
<% Else %>
<input type=hidden name=web_num_1 value='<% Response.Write web_var(web_num,1) %>'>
<input type=hidden name=web_num_2 value='<% Response.Write web_var(web_num,2) %>'>
<input type=hidden name=web_num_3 value='<% Response.Write web_var(web_num,3) %>'>
<input type=hidden name=web_num_4 value='<% Response.Write web_var(web_num,4) %>'>
<input type=hidden name=web_num_5 value='<% Response.Write web_var(web_num,5) %>'>
<input type=hidden name=web_num_6 value='<% Response.Write web_var(web_num,6) %>'>
<input type=hidden name=web_num_7 value='<% Response.Write web_var(web_num,7) %>'>
<input type=hidden name=web_num_8 value='<% Response.Write web_var(web_num,8) %>'>
<input type=hidden name=web_num_9 value='<% Response.Write web_var(web_num,9) %>'>
<input type=hidden name=web_num_10 value='<% Response.Write web_var(web_num,10) %>'>
<input type=hidden name=web_num_11 value='<% Response.Write web_var(web_num,11) %>'>
<input type=hidden name=web_num_12 value='<% Response.Write web_var(web_num,12) %>'>
<input type=hidden name=web_num_13 value='<% Response.Write web_var(web_num,13) %>'>
<input type=hidden name=web_num_14 value='<% Response.Write web_var(web_num,14) %>'>
<input type=hidden name=web_num_15 value='<% Response.Write web_var(web_num,15) %>'>
<input type=hidden name=web_num_16 value='<% Response.Write web_var(web_num,16) %>'>
<%
    End If

    If action = "menu" Then
        tt   = 1
        tdim = Split(web_menu,"|")

        For i = 0 To UBound(tdim) %>
<tr>
<td>网站菜单 <% Response.Write i + 1 %>：</td>
<td colspan=2><input type=text name=web_menu_<% Response.Write i + 1 %> value='<% Response.Write tdim(i) %>' size=40 maxlength=20></td>
</tr>
<%
        Next

        Erase tdim %>
<input type=hidden name=web_menu_num value='<% Response.Write i %>'>
<tr>
<td>增加新菜单：</td>
<td colspan=2><input type=text name=web_menu_new value='' size=40 maxlength=20>&nbsp;例：abc:文章学习</td>
</tr>
<tr>
<td colspan=3 class=htd><font class=red>网站菜单修改说明：</font>你可以添加、修改、删除网站菜单，但文本框中只能是以“abc:文章学习”的形式（栏目名:菜单名）存在，否则将出错！！！<br>
<font class=red_3>添加菜单</font>即在“增加新菜单”中按规则加入要新添的菜单，只能一个一个的增加；<br>
<font class=red_3>修改菜单</font>可修改所有的，如要排序，只要将其内容互换即可，但栏目名和菜单名要同时互换；<br>
<font class=red_3>删除菜单</font>可一次删除一个或多个，方法为将要删除的栏目菜单框里的内容清空，再将栏目数较大的向较小的里转移。</td>
</tr>
<% Else %>
<input type=hidden name=web_menu value='<% Response.Write web_menu %>'>
<%
    End If

    If action = "color" Then
        tt = 1 %>
<tr>
<td>网站背景：</td>
<td><input type=text name=web_color_1 value='<% Response.Write web_var(web_color,1) %>' size=20 maxlength=7></td>
<td bgcolor=<% Response.Write web_var(web_color,1) %>>&nbsp;</td>
</tr>
<tr>
<td>表格色一：</td>
<td><input type=text name=web_color_2 value='<% Response.Write web_var(web_color,2) %>' size=20 maxlength=7></td>
<td bgcolor=<% Response.Write web_var(web_color,2) %>>&nbsp;</td>
</tr>
<tr>
<td>表格色二：</td>
<td><input type=text name=web_color_3 value='<% Response.Write web_var(web_color,3) %>' size=20 maxlength=7></td>
<td bgcolor=<% Response.Write web_var(web_color,3) %>>&nbsp;</td>
</tr>
<tr>
<td>表格色三：</td>
<td><input type=text name=web_color_4 value='<% Response.Write web_var(web_color,4) %>' size=20 maxlength=7></td>
<td bgcolor=<% Response.Write web_var(web_color,4) %>>&nbsp;</td>
</tr>
<tr>
<td>表格色四：</td>
<td><input type=text name=web_color_5 value='<% Response.Write web_var(web_color,5) %>' size=20 maxlength=7></td>
<td bgcolor=<% Response.Write web_var(web_color,5) %>>&nbsp;</td>
</tr>
<tr>
<td>左边背景：</td>
<td><input type=text name=web_color_6 value='<% Response.Write web_var(web_color,6) %>' size=20 maxlength=7></td>
<td bgcolor=<% Response.Write web_var(web_color,6) %>>&nbsp;</td>
</tr>
<tr>
<td>主字体色：</td>
<td><input type=text name=web_color_7 value='<% Response.Write web_var(web_color,7) %>' size=20 maxlength=7></td>
<td bgcolor=<% Response.Write web_var(web_color,7) %>>&nbsp;</td>
</tr>
<tr>
<td>突出字体一：</td>
<td><input type=text name=web_color_8 value='<% Response.Write web_var(web_color,8) %>' size=20 maxlength=7></td>
<td bgcolor=<% Response.Write web_var(web_color,8) %>>&nbsp;</td>
</tr>
<tr>
<td>淡色字体：</td>
<td><input type=text name=web_color_9 value='<% Response.Write web_var(web_color,9) %>' size=20 maxlength=7></td>
<td bgcolor=<% Response.Write web_var(web_color,9) %>>&nbsp;</td>
</tr>
<tr>
<td>红色字体一：</td>
<td><input type=text name=web_color_10 value='<% Response.Write web_var(web_color,10) %>' size=20 maxlength=7></td>
<td bgcolor=<% Response.Write web_var(web_color,10) %>>&nbsp;</td>
</tr>
<tr>
<td>红色字体二：</td>
<td><input type=text name=web_color_11 value='<% Response.Write web_var(web_color,11) %>' size=20 maxlength=7></td>
<td bgcolor=<% Response.Write web_var(web_color,11) %>>&nbsp;</td>
</tr>
<tr>
<td>红色字体三：</td>
<td><input type=text name=web_color_12 value='<% Response.Write web_var(web_color,12) %>' size=20 maxlength=7></td>
<td bgcolor=<% Response.Write web_var(web_color,12) %>>&nbsp;</td>
</tr>
<tr>
<td>网站字体：</td>
<td colspan=2><input type=text name=web_font_family value='<% Response.Write web_font_family %>' size=60 maxlength=100></td>
</tr>
<tr>
<td>字体大小：</td>
<td colspan=2><input type=text name=web_font_size value='<% Response.Write web_font_size %>' size=20 maxlength=10></td>
</tr>
<% Else %>
<input type=hidden name=web_color_1 value='<% Response.Write web_var(web_color,1) %>'>
<input type=hidden name=web_color_2 value='<% Response.Write web_var(web_color,2) %>'>
<input type=hidden name=web_color_3 value='<% Response.Write web_var(web_color,3) %>'>
<input type=hidden name=web_color_4 value='<% Response.Write web_var(web_color,4) %>'>
<input type=hidden name=web_color_5 value='<% Response.Write web_var(web_color,5) %>'>
<input type=hidden name=web_color_6 value='<% Response.Write web_var(web_color,6) %>'>
<input type=hidden name=web_color_7 value='<% Response.Write web_var(web_color,7) %>'>
<input type=hidden name=web_color_8 value='<% Response.Write web_var(web_color,8) %>'>
<input type=hidden name=web_color_9 value='<% Response.Write web_var(web_color,9) %>'>
<input type=hidden name=web_color_10 value='<% Response.Write web_var(web_color,10) %>'>
<input type=hidden name=web_color_11 value='<% Response.Write web_var(web_color,11) %>'>
<input type=hidden name=web_color_12 value='<% Response.Write web_var(web_color,12) %>'>
<input type=hidden name=web_font_family value='<% Response.Write web_font_family %>'>
<input type=hidden name=web_font_size value='<% Response.Write web_font_size %>'>
<%
    End If

    If action = "down_up" Then
        tt = 1 %>
<tr>
<td>软件图片宽：</td>
<td><input type=text name=web_down_1 value='<% Response.Write web_var(web_down,1) %>' size=20 maxlength=10>&nbsp;像素</td>
<td class=gray>&nbsp;软件图片的宽度</td>
</tr>
<tr>
<td>软件图片高：</td>
<td><input type=text name=web_down_2 value='<% Response.Write web_var(web_down,2) %>' size=20 maxlength=10>&nbsp;像素</td>
<td class=gray>&nbsp;软件图片的高度</td>
</tr>
<tr>
<td>软件目录：</td>
<td><input type=text name=web_down_5 value='<% Response.Write web_var(web_down,5) %>' size=20 maxlength=20></td>
<td class=gray>&nbsp;</td>
</tr>
<tr>
<td>运行环境：</td>
<td colspan=2><input type=text name=web_down_3 value='<% Response.Write web_var(web_down,3) %>' size=66 maxlength=100></td>
</tr>
<tr>
<td>授权类型：</td>
<td colspan=2><input type=text name=web_down_4 value='<% Response.Write web_var(web_down,4) %>' size=66 maxlength=100></td>
</tr>
<tr>
<td>上传路径：</td>
<td><input type=text name=web_upload_1 value='<% Response.Write web_var(web_upload,1) %>' size=20 maxlength=20></td>
<td class=gray>&nbsp;为防有错，尽量不要修改此项</td>
</tr>
<tr>
<td>文件类型：</td>
<td><input type=text name=web_upload_2 value='<% Response.Write web_var(web_upload,2) %>' size=35 maxlength=50></td>
<td class=gray>&nbsp;多个类型用“,”分开</td>
</tr>
<tr>
<td>文件大小：</td>
<td><input type=text name=web_upload_3 value='<% Response.Write web_var(web_upload,3) %>' size=20 maxlength=10></td>
<td class=gray>&nbsp;单位为KB，最好小于500</td>
</tr>
<td>下载目录：</td>
<td><input type=text name=web_upload_3 value='<% Response.Write web_var(web_lct,1) %>' size=20 maxlength=10></td>
<td class=gray>&nbsp;</td>
</tr>
<td>软件目录：</td>
<td><input type=text name=web_upload_3 value='<% Response.Write web_var(web_lct,2) %>' size=20 maxlength=10></td>
<td class=gray>&nbsp;</td>
</tr>
<td>视频目录：</td>
<td><input type=text name=web_upload_3 value='<% Response.Write web_var(web_lct,3) %>' size=20 maxlength=10></td>
<td class=gray>&nbsp;</td>
</tr>
<td>Flash目录：</td>
<td><input type=text name=web_upload_3 value='<% Response.Write web_var(web_lct,4) %>' size=20 maxlength=10></td>
<td class=gray>&nbsp;</td>
</tr>
<td>壁纸目录：</td>
<td><input type=text name=web_upload_3 value='<% Response.Write web_var(web_lct,5) %>' size=20 maxlength=10></td>
<td class=gray>&nbsp;</td>
</tr>
<% Else %>
<input type=hidden name=web_down_1 value='<% Response.Write web_var(web_down,1) %>'>
<input type=hidden name=web_down_2 value='<% Response.Write web_var(web_down,2) %>'>
<input type=hidden name=web_down_5 value='<% Response.Write web_var(web_down,5) %>'>
<input type=hidden name=web_down_3 value='<% Response.Write web_var(web_down,3) %>'>
<input type=hidden name=web_down_4 value='<% Response.Write web_var(web_down,4) %>'>
<input type=hidden name=web_upload_1 value='<% Response.Write web_var(web_upload,1) %>'>
<input type=hidden name=web_upload_2 value='<% Response.Write web_var(web_upload,2) %>'>
<input type=hidden name=web_upload_3 value='<% Response.Write web_var(web_upload,3) %>'>
<%
    End If

    If action = "info" Then
        tt = 1 %>
<tr><td colspan=3 class=red_3>&nbsp;过滤字符设置</td></tr>
<tr>
<td>非法字符：</td>
<td colspan=2><input type=text name=web_safety_1 value='<% Response.Write Replace(web_var(web_safety,1),"'","") %>' size=30 maxlength=100>&nbsp;&nbsp;单引号(')和双引号(")已被系统过滤</td>
</tr>
<tr>
<td>密码允许：</td>
<td colspan=2><input type=text name=web_safety_2 value='<% Response.Write web_var(web_safety,2) %>' size=66 maxlength=200></td>
</tr>
<tr>
<td>注册禁用：</td>
<td colspan=2><input type=text name=web_safety_3 value='<% Response.Write web_var(web_safety,3) %>' size=66 maxlength=200></td>
</tr>
<tr>
<td>不健康字符：</td>
<td colspan=2><input type=text name=web_safety_4 value='<% Response.Write web_var(web_safety,4) %>' size=66 maxlength=200></td>
</tr>
<tr><td colspan=3 class=red_3>&nbsp;信息提示设置</td></tr>
<tr>
<td>外部提交：</td>
<td colspan=2><input type=text name=web_error_1 value='<% Response.Write web_var(web_error,1) %>' size=66 maxlength=200></td>
</tr>
<tr>
<td>未注册登陆：</td>
<td colspan=2><input type=text name=web_error_2 value='<% Response.Write web_var(web_error,2) %>' size=66 maxlength=200></td>
</tr>
<tr>
<td>支持信息：</td>
<td colspan=2><input type=text name=web_error_3 value='<% Response.Write web_var(web_error,3) %>' size=66 maxlength=200></td>
</tr>
<tr>
<td>网站底部：</td>
<td colspan=2><input type=text name=web_error_4 value='<% Response.Write web_var(web_error,4) %>' size=66 maxlength=200></td>
</tr>
<% Else %>
<input type=hidden name=web_safety_1 value='<% Response.Write Replace(web_var(web_safety,1),"'","") %>'>
<input type=hidden name=web_safety_2 value='<% Response.Write web_var(web_safety,2) %>'>
<input type=hidden name=web_safety_3 value='<% Response.Write web_var(web_safety,3) %>'>
<input type=hidden name=web_safety_4 value='<% Response.Write web_var(web_safety,4) %>'>
<input type=hidden name=web_error_1 value='<% Response.Write web_var(web_error,1) %>'>
<input type=hidden name=web_error_2 value='<% Response.Write web_var(web_error,2) %>'>
<input type=hidden name=web_error_3 value='<% Response.Write web_var(web_error,3) %>'>
<input type=hidden name=web_error_4 value='<% Response.Write web_var(web_error,4) %>'>
<%
    End If

    If action = "user" Then
        tt   = 1
        tdim = Split(user_power,"|")

        For i = 0 To UBound(tdim) %>
<tr>
<td>用户组 <% Response.Write i + 1 %>：</td>
<td><input type=text name=user_power_<% Response.Write i + 1 %> value='<% Response.Write tdim(i) %>' size=30 maxlength=20></td>
<td class=gray>&nbsp;</td>
</tr>
<%
        Next

        Erase tdim %>
<input type=hidden name=user_power_num value='<% Response.Write i %>'>
<tr>
<td>新用户组：</td>
<td><input type=text name=user_power_new value='' size=30 maxlength=20></td>
<td class=gray>&nbsp;例：huser:高级用户</td>
</tr>
<tr>
<td colspan=3 class=htd><font class=red>用户组修改说明：</font>你可以添加、修改、删除用户组，但文本框中只能是以“huser:高级用户”的形式（组别:用户组名）存在，否则将出错！！！
<font class=red>在程序中，用户组数越小，权限越大，前三个用户组别请勿私自修改，但修改其用户组名可以！当网站正式运行后，请勿再修改用户组别，以防部分用户将会不能被程序正确实别！</font><br>
<font class=red_3>添加用户组</font>即在“增加新用户组”中按规则加入要新添的用户组，只能单个的增加；<br>
<font class=red_3>修改用户组</font>可修改所有的，如要排序，只要将其内容互换即可，但用户组别和用户组名要同时互换；<br>
<font class=red_3>删除用户组</font>可一次删除一个或多个，方法为将要删除的用户组框里的内容清空，再将栏目数较大的向较小的里转移。</td>
</tr>
<% Else %>
<input type=hidden name=user_power value='<% Response.Write user_power %>'>
<%
    End If

    If action = "grade" Then
        tt   = 1
        tdim = Split(user_grade,"|")

        For i = 0 To UBound(tdim) %>
<tr>
<td>用户等级 <% Response.Write i %>：</td>
<td><input type=text name=user_grade_<% Response.Write i + 1 %> value='<% Response.Write tdim(i) %>' size=30 maxlength=20></td>
<td class=gray>&nbsp;<img src='images/star/star_<% Response.Write i %>.gif' border=0></td>
</tr>
<%
        Next

        Erase tdim %>
<input type=hidden name=user_grade_num value='<% Response.Write i %>'>
<tr>
<td>新用户等级：</td>
<td><input type=text name=user_grade_new value='' size=30 maxlength=20></td>
<td class=gray>&nbsp;例：10000:超级</td>
</tr>
<tr>
<td colspan=3 class=htd><font class=red>用户等级修改说明：</font>你可以添加、修改、删除用户等级，但文本框中只能是以“10000:超级”的形式（所需积分:等级名称）存在，否则将出错！！！<br>
<font class=red_3>添加用户等级</font>即在“增加新用户组”中按规则加入要新添的用户等级，只能单个的增加；<br>
<font class=red_3>修改用户等级</font>可修改所有的，如要排序，只要将其内容互换即可，但所需积分和等级名称要同时互换；<br>
<font class=red_3>删除用户等级</font>可一次删除一个或多个，方法为将要删除的用户等级框里的内容清空，再将栏目数较大的向较小的里转移。</td>
</tr>
<% Else %>
<input type=hidden name=user_grade value='<% Response.Write user_grade %>'>
<%
    End If

    If action = "forum" Then
        tt     = 1
        udim   = Split(user_power,"|")
        tdim   = Split(forum_type,"|")

        For i = 0 To UBound(tdim)
            t2 = Left(tdim(i),InStr(tdim(i),":") - 1) %>
<tr>
<td>论坛分类 <% Response.Write i + 1 %>：</td>
<td><input type=text name=forum_type_<% Response.Write i + 1 %>_2 value='<% Response.Write Right(tdim(i),Len(tdim(i)) - InStr(tdim(i),":")) %>' size=20 maxlength=20></td>
<td class=gray>&nbsp;</td>
</tr>
<tr><td class=gray>论坛权限</td><td colspan=2><%

            For j = 0 To UBound(udim)
                Response.Write vbcrlf & "<input type=checkbox name=forum_type_" & i + 1 & "_1 value='" & j + 1 & "' class=bg_1"
                If InStr(1,"." & t2 & ".","." & j + 1 & ".") > 0 Then Response.Write " checked"
                Response.Write ">" & Right(udim(j),Len(udim(j)) - InStr(udim(j),":"))
            Next %><input type=checkbox name=forum_type_<% Response.Write i + 1 %>_1 value='0' class=bg_1<% If InStr(1,"." & t2 & ".",".0.") > 0 Then Response.Write " checked" %>>游客</td></tr><%
        Next

        Erase tdim %>
<input type=hidden name=forum_type_num value='<% Response.Write i %>'>
<tr>
<td>新论坛分类：</td>
<td><input type=text name=forum_type_new_2 value='' size=30 maxlength=20></td>
<td class=gray>&nbsp;例：精华论坛</td>
</tr>
<tr><td class=gray>论坛权限</td><td colspan=2><%

        For j = 0 To UBound(udim)
            Response.Write vbcrlf & "<input type=checkbox name=forum_type_new_1 value='" & j + 1 & "' class=bg_1>" & Right(udim(j),Len(udim(j)) - InStr(udim(j),":"))
        Next %><input type=checkbox name=forum_type_new_1 value='0' class=bg_1>游客</td></tr>
<tr>
<td colspan=3 class=htd><font class=red>论坛分类修改说明：</font>你可以添加、修改、删除论坛分类，但文本框中只能是以“文章学习”的形式（论坛分类名称）存在，否则将出错！！！
<font class=red>如用户组有删除或修改其权限后，请重新分配论坛权限！如论坛分类及权限有修改，请进入论坛管理以重新分配论坛类别！</font><br>
<font class=red_3>添加论坛分类</font>即在“新论坛分类”中按规则加入要新添的论坛分类，只能单个的增加；<br>
<font class=red_3>修改论坛分类</font>可修改所有的，排序对论坛的权限没有影响；<br>
<font class=red_3>删除论坛分类</font>可一次删除一个或多个，方法为将要删除的论坛分类框里的内容删除即可。</td>
</tr>
<%
        Erase udim
    Else %>
<input type=hidden name=forum_type value='<% Response.Write forum_type %>'>
<%
    End If

    If tt = 0 Then %>
<tr><td colspan=3 align=center height=150 class=htd><font class=red>请先选择配置修改的类型！建议在修改配置文件前先<a href='javascript:Do_b_data();'>备份配置</a>。<br>如果在修改配置时出现了错误，您可以试着<a href='javascript:Do_h_data();'>还原配置</a>。<br>如果因为您的一时不小心，将配置文件修改错误，请尽快用FTP连上并用可用的配置文件覆盖错误的文件（include/common.asp）</font></td></tr>
<% Else %>
<tr><td colspan=3 align=center height=30><input type=submit value='提 交 修 改'>
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
<input type=reset value=' 重 置 '>
</td></tr>
<% End If %>
</form>
</table>
<%

End Sub

Sub config_chk()
    Dim t1
    Dim t2
    Dim tn
    web_font_family = code_config(Request.form("web_font_family"),0):web_font_size = code_config(Request.form("web_font_size"),0)
    web_config      = code_config(Request.form("web_config_1"),3) & "|" & code_config(Request.form("web_config_2"),3) & "|" & _
    code_config(Request.form("web_config_3"),3) & "|" & code_config(Request.form("web_config_4"),3) & "|" & _
    code_config(Request.form("web_config_5"),3) & "|" & code_config(Request.form("web_config_6"),3) & "|" & _
    code_config(Request.form("web_config_7"),3) & "|" & code_config(Request.form("web_config_8"),3) & "|" & _
    Trim(Request.form("web_config_9_1")) & Trim(Request.form("web_config_9_2")) & _
    Trim(Request.form("web_config_9_3")) & Trim(Request.form("web_config_9_4")) & _
    Trim(Request.form("web_config_9_5"))
    web_cookies = code_config(Request.form("web_cookies"),3):web_login = Trim(Request.form("web_login"))
    web_setup   = Trim(Request.form("web_setup_1")) & Trim(Request.form("web_setup_2")) & Trim(Request.form("web_setup_3")) & _
    Trim(Request.form("web_setup_4")) & Trim(Request.form("web_setup_5")) & Trim(Request.form("web_setup_6")) & _
    Trim(Request.form("web_setup_7"))
    web_num = code_num(Trim(Request.form("web_num_1")),20) & "|" & code_num(Trim(Request.form("web_num_2")),20) & "|" & _
    code_num(Trim(Request.form("web_num_3")),10) & "|" & code_num(Trim(Request.form("web_num_4")),5) & "|" & _
    code_num(Trim(Request.form("web_num_5")),3) & "|" & code_num(Trim(Request.form("web_num_6")),12) & "|" & _
    code_num(Trim(Request.form("web_num_7")),150) & "|" & code_num(Trim(Request.form("web_num_8")),112) & "|" & _
    code_num(Trim(Request.form("web_num_9")),430) & "|" & code_num(Trim(Request.form("web_num_10")),350) & "|" & _
    code_num(Trim(Request.form("web_num_11")),24) & "|" & code_num(Trim(Request.form("web_num_12")),16) & "|" & _
    code_num(Trim(Request.form("web_num_13")),20) & "|" & code_num(Trim(Request.form("web_num_14")),20) & "|" & _
    code_num(Trim(Request.form("web_num_15")),2) & "|" & code_num(Trim(Request.form("web_num_16")),15)

    If action = "menu" Then
        web_menu = "":tn = Int(Trim(Request.form("web_menu_num")))

        For i = 1 To tn
            t1   = code_config(Request.form("web_menu_" & i),4)
            If Len(t1) > 1 Then If InStr(1,t1,":") > 0 Then web_menu = web_menu & t1 & "|"
        Next

        t1 = code_config(Request.form("web_menu_new"),4)
        If Len(t1) > 1 Then If InStr(1,t1,":") > 0 Then web_menu = web_menu & t1
        If Right(web_menu,1) = "|" Then web_menu = Left(web_menu,Len(web_menu) - 1)
    Else
        web_menu = Trim(Request.form("web_menu"))
    End If

    web_color    = code_config(Request.form("web_color_1"),2) & "|" & code_config(Request.form("web_color_2"),2) & "|" & code_config(Request.form("web_color_3"),2) & "|" & _
    code_config(Request.form("web_color_4"),2) & "|" & code_config(Request.form("web_color_5"),2) & "|" & code_config(Request.form("web_color_6"),2) & "|" & _
    code_config(Request.form("web_color_7"),2) & "|" & code_config(Request.form("web_color_8"),2) & "|" & code_config(Request.form("web_color_9"),2) & "|" & _
    code_config(Request.form("web_color_10"),2) & "|" & code_config(Request.form("web_color_11"),2) & "|" & code_config(Request.form("web_color_12"),2)
    web_upload     = code_config(Request.form("web_upload_1"),2) & "|" & code_config(Request.form("web_upload_2"),2) & "|" & code_config(Request.form("web_upload_3"),2)
    web_safety     = Replace(Trim(Request.form("web_safety_1")),"""","""&chr(34)&""") & "'" & "|" & Trim(Request.form("web_safety_2")) & "|" & Trim(Request.form("web_safety_3")) & "|" & Trim(Request.form("web_safety_4"))
    web_error      = code_config(Request.form("web_error_1"),3) & "|" & code_config(Request.form("web_error_2"),3) & "|" & code_config(Request.form("web_error_3"),3) & "|" & code_config(Request.form("web_error_4"),3)
    web_news_art   = Trim(Request.form("web_news_art"))
    web_down       = code_num(Trim(Request.form("web_down_1")),95) & "|" & code_num(Trim(Request.form("web_down_2")),75) & "|" & _
    code_config(Request.form("web_down_3"),4) & "|" & code_config(Request.form("web_down_4"),4) & "|" & code_config(Request.form("web_down_5"),4)
    web_shop       = Trim(Request.form("web_shop"))
    web_stamp      = code_config(Request.form("web_stamp_1"),3) & "|" & code_config(Request.form("web_stamp_2"),3) & "|" & code_config(Request.form("web_stamp_3"),3)

    If action = "user" Then
        user_power = "":tn = Int(Trim(Request.form("user_power_num")))

        For i = 1 To tn
            t1     = code_config(Request.form("user_power_" & i),4)
            If Len(t1) > 1 Then If InStr(1,t1,":") > 0 Then user_power = user_power & t1 & "|"
        Next

        t1 = code_config(Request.form("user_power_new"),4)
        If Len(t1) > 1 Then If InStr(1,t1,":") > 0 Then user_power = user_power & t1
        If Right(user_power,1) = "|" Then user_power = Left(user_power,Len(user_power) - 1)
    Else
        user_power = Trim(Request.form("user_power"))
    End If

    If action = "grade" Then
        user_grade = "":tn = Int(Trim(Request.form("user_grade_num")))

        For i = 1 To tn
            t1     = code_config(Request.form("user_grade_" & i),4)
            If Len(t1) > 1 Then If InStr(1,t1,":") > 0 Then user_grade = user_grade & t1 & "|"
        Next

        t1 = code_config(Request.form("user_grade_new"),4)
        If Len(t1) > 1 Then If InStr(1,t1,":") > 0 Then user_grade = user_grade & t1
        If Right(user_grade,1) = "|" Then user_grade = Left(user_grade,Len(user_grade) - 1)
    Else
        user_grade = Trim(Request.form("user_grade"))
    End If

    If action = "forum" Then
        forum_type = "":tn = Int(Trim(Request.form("forum_type_num")))

        For i = 1 To tn
            t1     = Replace(Trim(Request.form("forum_type_" & i & "_1"))," ","")
            t2     = code_config(Request.form("forum_type_" & i & "_2"),2)
            If Len(t1) > 0 And Len(t2) > 0 Then t1 = Replace(t1,",","."):forum_type = forum_type & t1 & ":" & t2 & "|"
        Next

        t1 = Replace(Trim(Request.form("forum_type_new_1"))," ","")
        t2 = code_config(Request.form("forum_type_new_2"),2)
        If Len(t1) > 0 And Len(t2) > 0 Then t1 = Replace(t1,",","."):forum_type = forum_type & t1 & ":" & t2
        If Right(forum_type,1) = "|" Then forum_type = Left(forum_type,Len(forum_type) - 1)
    Else
        forum_type = Trim(Request.form("forum_type"))
    End If

    Call config_file()
    If action = "color" Then Call config_css():Call config_mouse_on_title()
    Response.Write "<script language=javascript>alert(""配置修改成功！"");</script>"
    Call config_main()
End Sub

Sub config_bh(bht)
    Dim vv
    Dim filetype
    Dim file_name1
    Dim file_name2
    Dim filetemp
    Dim fileos
    Dim filepath:filetype = ""

    If bht = "h" Then
        file_name1 = "include/back_common.asp"
        file_name2 = "include/common.asp"
        vv         = "还原"
    Else
        file_name1 = "include/common.asp"
        file_name2 = "include/back_common.asp"
        vv         = "备份"
    End If

    Set fileos   = CreateObject("Scripting.FileSystemObject")
    filepath     = Server.mappath(file_name1)
    Set filetemp = fileos.OpenTextFile(filepath,1,True)
    filetype     = filetemp.ReadAll
    filetemp.Close
    filepath     = Server.mappath(file_name2)
    Set filetemp = fileos.CreateTextFile(filepath,True)
    filetemp.WriteLine( filetype )
    filetemp.Close
    Set filetemp = Nothing
    Set fileos   = Nothing
    Response.Write "<script language=javascript>alert(""" & vv & " 网站配置成功！"");</script>"
End Sub

Function code_num(strers,cnum)
    Dim strer:strer = Trim(strers)
    If Not(IsNumeric(strer)) Then strer = cnum
    If Int(strer) < 1 Then strer = cnum
    code_num = strer
End Function

Function code_config(strers,ct)
    Dim strer:strer = Trim(strers)
    If IsNull(strer) Or strer = "" Then code_config = "":Exit Function

    Select Case ct
        Case 1
            strer = Replace(strer,"'","&#39;")
            strer = Replace(strer,Chr(34),"&quot;")
        Case 2
            strer = Replace(strer,Chr(34),"")
            strer = Replace(strer,"'","")
            strer = Replace(strer,":","")
            strer = Replace(strer,"|","")
        Case 3
            strer = Replace(strer,"'","&#39;")
            strer = Replace(strer,Chr(34),"&quot;")
            strer = Replace(strer,"|","")
        Case 4
            strer = Replace(strer,Chr(34),"")
            strer = Replace(strer,"'","")
            strer = Replace(strer,"|","")
        Case Else
            strer = Replace(strer,"'","")
            strer = Replace(strer,Chr(34),"")
    End Select

    code_config = strer
End Function %>
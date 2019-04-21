<%
'*******************************************************************
'
'                     Beyondest.Com v3.6.1
' 
'           http://beyondest.com
' 
'*******************************************************************

if action="config" then
  tt=1
%>
<tr>
<td>网站名称：</td>
<td><input type=text name=web_config_1 value='<%response.write web_var(web_config,1)%>' size=38 maxlength=20></td>
<td class=gray>&nbsp;</td>
</tr>
<tr>
<td>网站地址：</td>
<td><input type=text name=web_config_2 value='<%response.write web_var(web_config,2)%>' size=38 maxlength=50></td>
<td class=gray>&nbsp;</td>
</tr>
<tr>
<td>超级管理员：</td>
<td><input type=text name=web_config_3 value='<%response.write web_var(web_config,3)%>' size=20 maxlength=20></td>
<td class=gray>&nbsp;</td>
</tr>
<tr>
<td>所在目录：</td>
<td><input type=text name=web_config_4 value='<%response.write web_var(web_config,4)%>' size=20 maxlength=20></td>
<td class=gray>&nbsp;</td>
</tr>
<tr>
<td>网站SKIN：</td>
<td><input type=text name=web_config_5 value='<%response.write web_var(web_config,5)%>' size=20 maxlength=20></td>
<td class=gray>&nbsp;</td>
</tr>
<tr>
<td>Cookies名称：</td>
<td><input type=text name=web_cookies value='<%response.write web_cookies%>' size=20 maxlength=20></td>
<td class=gray>&nbsp;</td>
</tr>
<tr>
<td>数据库名称：</td>
<td><input type=text name=web_config_6 value='<%response.write web_var(web_config,6)%>' size=38 maxlength=50></td>
<td class=gray>&nbsp;</td>
</tr>
<tr>
<td>网站背景：</td>
<td><input type=text name=web_config_7 value='<%response.write web_var(web_config,7)%>' size=20 maxlength=20></td>
<td background='images/<%response.write web_var(web_config,7)%>.gif'>&nbsp;</td>
</tr>
<tr>
<td>网站状态：</td>
<td><input type=radio name=web_login value='1'<% if int(web_login)=1 then response.write " checked" %> class=bg_1>&nbsp;开放&nbsp;<input type=radio name=web_login value='0'<% if int(web_login)<>1 then response.write " checked" %> class=bg_1>&nbsp;关闭</td>
<td class=gray>&nbsp;是否开放网站</td>
</tr>
<% t1=int(mid(web_setup,1,1)) %>
<tr>
<td>登陆浏览：</td>
<td colspan=2><input type=radio name=web_setup_1 value='1'<% if t1=1 then response.write " checked" %> class=bg_1>&nbsp;是&nbsp;&nbsp;&nbsp;<input type=radio name=web_setup_1 value='0'<% if t1<>1 then response.write " checked" %> class=bg_1>&nbsp;否&nbsp;&nbsp;&nbsp;&nbsp;<font class=gray>是否要登陆才可浏览文章或下载软件等</font></td>
</tr>
<%
t1=int(mid(web_setup,3,1))
if t1<>1 and t1<>2 then t1=0
%>
<tr>
<td>网站模式：</td>
<td colspan=2 class=gray><input type=radio name=web_setup_3 value='1'<% if t1=0 then response.write " checked" %> class=bg_1>&nbsp;<input type=text name=web_stamp_1 value='<%response.write web_var(web_stamp,1)%>' size=10 maxlength=10>&nbsp;&nbsp;&nbsp;本站的注册用户可以登陆，不记录在线列表</td>
<tr>
<td>&nbsp;</td>
<td colspan=2 class=gray><input type=radio name=web_setup_3 value='1'<% if t1=1 then response.write " checked" %> class=bg_1>&nbsp;<input type=text name=web_stamp_2 value='<%response.write web_var(web_stamp,2)%>' size=10 maxlength=10>&nbsp;&nbsp;&nbsp;所有登陆和浏览本站的人被并记录在线列表</td>
<tr>
<td>&nbsp;</td>
<td colspan=2 class=gray><input type=radio name=web_setup_3 value='2'<% if t1=2 then response.write " checked" %> class=bg_1>&nbsp;<input type=text name=web_stamp_3 value='<%response.write web_var(web_stamp,3)%>' size=10 maxlength=10>&nbsp;&nbsp;&nbsp;本站的注册用户可以登陆，并记录用户在线列表</td>
</tr>
<% t1=int(mid(web_setup,4,1)) %>
<tr>
<td>信息过滤：</td>
<td><input type=radio name=web_setup_4 value='1'<% if t1=1 then response.write " checked" %> class=bg_1>&nbsp;是&nbsp;&nbsp;&nbsp;<input type=radio name=web_setup_4 value='0'<% if t1<>1 then response.write " checked" %> class=bg_1>&nbsp;否</td>
<td class=gray>&nbsp;是否对论坛等进行字符过滤</td>
</tr>
<%
t1=int(mid(web_setup,5,1))
if t1<>0 and t1<>1 then t1=2
%>
<tr>
<td>显示 IP：</td>
<td><input type=radio name=web_setup_5 value='0'<% if t1=0 then response.write " checked" %> class=bg_1> 完全保密
<input type=radio name=web_setup_5 value='1'<% if t1=1 then response.write " checked" %> class=bg_1> 显示部分
<input type=radio name=web_setup_5 value='2'<% if t1=2 then response.write " checked" %> class=bg_1> 完全开放</td>
<td class=gray>&nbsp;对管理员总是完全开放</td>
</tr>
<% t1=int(mid(web_setup,6,1)) %>
<tr>
<td>版主显示：</td>
<td><input type=radio name=web_setup_6 value='1'<% if t1=1 then response.write " checked" %> class=bg_1>&nbsp;文字链接&nbsp;<input type=radio name=web_setup_6 value='0'<% if t1<>1 then response.write " checked" %> class=bg_1>&nbsp;下拉菜单框</td>
<td class=gray>&nbsp;论坛版主显示模式</td>
</tr>
<% else %>
<input type=hidden name=web_config_1 value='<%response.write web_var(web_config,1)%>'>
<input type=hidden name=web_config_2 value='<%response.write web_var(web_config,2)%>'>
<input type=hidden name=web_config_3 value='<%response.write web_var(web_config,3)%>'>
<input type=hidden name=web_config_4 value='<%response.write web_var(web_config,4)%>'>
<input type=hidden name=web_config_5 value='<%response.write web_var(web_config,5)%>'>
<input type=hidden name=web_cookies value='<%response.write web_cookies%>'>
<input type=hidden name=web_config_6 value='<%response.write web_var(web_config,6)%>'>
<input type=hidden name=web_config_7 value='<%response.write web_var(web_config,7)%>'>
<input type=hidden name=web_login value='<%response.write web_login%>'>
<input type=hidden name=web_setup_1 value='<%response.write mid(web_setup,1,1)%>'>
<input type=hidden name=web_setup_3 value='<%response.write mid(web_setup,3,1)%>'>
<input type=hidden name=web_stamp_1 value='<%response.write web_var(web_stamp,1)%>'>
<input type=hidden name=web_stamp_2 value='<%response.write web_var(web_stamp,2)%>'>
<input type=hidden name=web_stamp_3 value='<%response.write web_var(web_stamp,3)%>'>
<input type=hidden name=web_setup_4 value='<%response.write mid(web_setup,4,1)%>'>
<input type=hidden name=web_setup_5 value='<%response.write mid(web_setup,5,1)%>'>
<input type=hidden name=web_setup_6 value='<%response.write mid(web_setup,6,1)%>'>
<%
end if

if action="put" then
  tt=1
  t1=int(mid(web_var(web_config,9),1,1))
%><tr>
<td>发布新闻：</td>
<td><input type=radio name=web_config_9_1 value='1'<% if t1=1 then response.write " checked" %> class=bg_1>&nbsp;开放&nbsp;<input type=radio name=web_config_9_1 value='0'<% if t1=0 then response.write " checked" %> class=bg_1>&nbsp;关闭</td>
<td class=gray>&nbsp;</td>
</tr><% t1=int(mid(web_var(web_config,9),2,1)) %>
<tr>
<td>发表文章：</td>
<td><input type=radio name=web_config_9_2 value='1'<% if t1=1 then response.write " checked" %> class=bg_1>&nbsp;开放&nbsp;<input type=radio name=web_config_9_2 value='0'<% if t1=0 then response.write " checked" %> class=bg_1>&nbsp;关闭</td>
<td class=gray>&nbsp;</td>
</tr><% t1=int(mid(web_var(web_config,9),3,1)) %>
<tr>
<td>添加音乐：</td>
<td><input type=radio name=web_config_9_3 value='1'<% if t1=1 then response.write " checked" %> class=bg_1>&nbsp;开放&nbsp;<input type=radio name=web_config_9_3 value='0'<% if t1=0 then response.write " checked" %> class=bg_1>&nbsp;关闭</td>
<td class=gray>&nbsp;</td>
</tr><% t1=int(mid(web_var(web_config,9),4,1)) %>
<tr>
<td>上传贴图：</td>
<td><input type=radio name=web_config_9_4 value='1'<% if t1=1 then response.write " checked" %> class=bg_1>&nbsp;开放&nbsp;<input type=radio name=web_config_9_4 value='0'<% if t1=0 then response.write " checked" %> class=bg_1>&nbsp;关闭</td>
<td class=gray>&nbsp;</td>
</tr><% t1=int(mid(web_var(web_config,9),5,1)) %>
<tr>
<td>推荐网站：</td>
<td><input type=radio name=web_config_9_5 value='1'<% if t1=1 then response.write " checked" %> class=bg_1>&nbsp;开放&nbsp;<input type=radio name=web_config_9_5 value='0'<% if t1=0 then response.write " checked" %> class=bg_1>&nbsp;关闭</td>
<td class=gray>&nbsp;</td>
</tr>
<% else %>
<input type=hidden name=web_config_9_1 value='<%response.write mid(web_var(web_config,9),1,1)%>'>
<input type=hidden name=web_config_9_2 value='<%response.write mid(web_var(web_config,9),2,1)%>'>
<input type=hidden name=web_config_9_3 value='<%response.write mid(web_var(web_config,9),3,1)%>'>
<input type=hidden name=web_config_9_4 value='<%response.write mid(web_var(web_config,9),4,1)%>'>
<input type=hidden name=web_config_9_5 value='<%response.write mid(web_var(web_config,9),5,1)%>'>
<%
end if

if action="num" then
  tt=1
%>
<tr>
<td>用户名长度：</td>
<td><input type=text name=web_num_1 value='<%response.write web_var(web_num,1)%>' size=20 maxlength=20></td>
<td class=gray>&nbsp;</td>
</tr>
<tr>
<td>每页主题数：</td>
<td><input type=text name=web_num_2 value='<%response.write web_var(web_num,2)%>' size=20 maxlength=20></td>
<td class=gray>&nbsp;论坛、会员每页显示等</td>
</tr>
<tr>
<td>每页显示数：</td>
<td><input type=text name=web_num_3 value='<%response.write web_var(web_num,3)%>' size=20 maxlength=20></td>
<td class=gray>&nbsp;查看贴子内容等</td>
</tr>
<tr>
<td>每页留言数：</td>
<td><input type=text name=web_num_4 value='<%response.write web_var(web_num,4)%>' size=20 maxlength=20></td>
<td class=gray>&nbsp;每页留言数目等</td>
</tr>
<tr>
<td>自动返回：</td>
<td><input type=text name=web_num_5 value='<%response.write web_var(web_num,5)%>' size=20 maxlength=20></td>
<td class=gray>&nbsp;自动返回时间，单位为秒</td>
</tr>
<tr>
<td>主题长度：</td>
<td><input type=text name=web_num_6 value='<%response.write web_var(web_num,6)%>' size=20 maxlength=20>&nbsp;KB</td>
<td class=gray>&nbsp;论坛发贴主题等长度</td>
</tr>
<tr>
<td>图片宽度：</td>
<td><input type=text name=web_num_7 value='<% response.write web_var(web_num,7) %>' size=20 maxlength=20>&nbsp;像素</td>
<td class=gray>&nbsp;下载、图库等图片显示宽度</td>
</tr>
<tr>
<td>图片高度：</td>
<td><input type=text name=web_num_8 value='<% response.write web_var(web_num,8) %>' size=20 maxlength=20>&nbsp;像素</td>
<td class=gray>&nbsp;下载、图库等图片显示高度</td>
</tr>
<tr>
<td>最大宽度：</td>
<td><input type=text name=web_num_9 value='<% response.write web_var(web_num,9) %>' size=20 maxlength=20>&nbsp;像素</td>
<td class=gray>&nbsp;贴图、FLASH等最大宽度</td>
</tr>
<tr>
<td>最大高度：</td>
<td><input type=text name=web_num_10 value='<% response.write web_var(web_num,10) %>' size=20 maxlength=20>&nbsp;像素</td>
<td class=gray>&nbsp;贴图、FLASH等最大高度</td>
</tr>
<tr>
<td>用户头像数：</td>
<td><input type=text name=web_num_11 value='<% response.write web_var(web_num,11) %>' size=20 maxlength=20></td>
<td class=gray>&nbsp;用户头像的总数</td>
</tr>
<% else %>
<input type=hidden name=web_num_1 value='<%response.write web_var(web_num,1)%>'>
<input type=hidden name=web_num_2 value='<%response.write web_var(web_num,2)%>'>
<input type=hidden name=web_num_3 value='<%response.write web_var(web_num,3)%>'>
<input type=hidden name=web_num_4 value='<%response.write web_var(web_num,4)%>'>
<input type=hidden name=web_num_5 value='<%response.write web_var(web_num,5)%>'>
<input type=hidden name=web_num_6 value='<%response.write web_var(web_num,6)%>'>
<input type=hidden name=web_num_7 value='<%response.write web_var(web_num,7)%>'>
<input type=hidden name=web_num_8 value='<%response.write web_var(web_num,8)%>'>
<input type=hidden name=web_num_9 value='<%response.write web_var(web_num,9)%>'>
<input type=hidden name=web_num_10 value='<%response.write web_var(web_num,10)%>'>
<input type=hidden name=web_num_11 value='<%response.write web_var(web_num,11)%>'>
<%
end if

if action="menu" then
  tt=1
  tdim=split(web_menu,"|")
  for i=0 to ubound(tdim)
%>
<tr>
<td>网站菜单 <%response.write i+1%>：</td>
<td colspan=2><input type=text name=web_menu_<%response.write i+1%> value='<%response.write tdim(i)%>' size=40 maxlength=20></td>
</tr>
<%
  next
  erase tdim
%>
<input type=hidden name=web_menu_num value='<%response.write i%>'>
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
<% else %>
<input type=hidden name=web_menu value='<%response.write web_menu%>'>
<%
end if

if action="color" then
  tt=1
%>
<tr>
<td>网站背景：</td>
<td><input type=text name=web_color_1 value='<%response.write web_var(web_color,1)%>' size=20 maxlength=7></td>
<td bgcolor=<%response.write web_var(web_color,1)%>>&nbsp;</td>
</tr>
<tr>
<td>表格色一：</td>
<td><input type=text name=web_color_2 value='<%response.write web_var(web_color,2)%>' size=20 maxlength=7></td>
<td bgcolor=<%response.write web_var(web_color,2)%>>&nbsp;</td>
</tr>
<tr>
<td>表格色二：</td>
<td><input type=text name=web_color_3 value='<%response.write web_var(web_color,3)%>' size=20 maxlength=7></td>
<td bgcolor=<%response.write web_var(web_color,3)%>>&nbsp;</td>
</tr>
<tr>
<td>表格色三：</td>
<td><input type=text name=web_color_4 value='<%response.write web_var(web_color,4)%>' size=20 maxlength=7></td>
<td bgcolor=<%response.write web_var(web_color,4)%>>&nbsp;</td>
</tr>
<tr>
<td>表格色四：</td>
<td><input type=text name=web_color_5 value='<%response.write web_var(web_color,5)%>' size=20 maxlength=7></td>
<td bgcolor=<%response.write web_var(web_color,5)%>>&nbsp;</td>
</tr>
<tr>
<td>左边背景：</td>
<td><input type=text name=web_color_6 value='<%response.write web_var(web_color,6)%>' size=20 maxlength=7></td>
<td bgcolor=<%response.write web_var(web_color,6)%>>&nbsp;</td>
</tr>
<tr>
<td>主字体色：</td>
<td><input type=text name=web_color_7 value='<%response.write web_var(web_color,7)%>' size=20 maxlength=7></td>
<td bgcolor=<%response.write web_var(web_color,7)%>>&nbsp;</td>
</tr>
<tr>
<td>突出字体一：</td>
<td><input type=text name=web_color_8 value='<%response.write web_var(web_color,8)%>' size=20 maxlength=7></td>
<td bgcolor=<%response.write web_var(web_color,8)%>>&nbsp;</td>
</tr>
<tr>
<td>淡色字体：</td>
<td><input type=text name=web_color_9 value='<%response.write web_var(web_color,9)%>' size=20 maxlength=7></td>
<td bgcolor=<%response.write web_var(web_color,9)%>>&nbsp;</td>
</tr>
<tr>
<td>红色字体一：</td>
<td><input type=text name=web_color_10 value='<%response.write web_var(web_color,10)%>' size=20 maxlength=7></td>
<td bgcolor=<%response.write web_var(web_color,10)%>>&nbsp;</td>
</tr>
<tr>
<td>红色字体二：</td>
<td><input type=text name=web_color_11 value='<%response.write web_var(web_color,11)%>' size=20 maxlength=7></td>
<td bgcolor=<%response.write web_var(web_color,11)%>>&nbsp;</td>
</tr>
<tr>
<td>红色字体三：</td>
<td><input type=text name=web_color_12 value='<%response.write web_var(web_color,12)%>' size=20 maxlength=7></td>
<td bgcolor=<%response.write web_var(web_color,12)%>>&nbsp;</td>
</tr>
<% else %>
<input type=hidden name=web_color_1 value='<%response.write web_var(web_color,1)%>'>
<input type=hidden name=web_color_2 value='<%response.write web_var(web_color,2)%>'>
<input type=hidden name=web_color_3 value='<%response.write web_var(web_color,3)%>'>
<input type=hidden name=web_color_4 value='<%response.write web_var(web_color,4)%>'>
<input type=hidden name=web_color_5 value='<%response.write web_var(web_color,5)%>'>
<input type=hidden name=web_color_6 value='<%response.write web_var(web_color,6)%>'>
<input type=hidden name=web_color_7 value='<%response.write web_var(web_color,7)%>'>
<input type=hidden name=web_color_8 value='<%response.write web_var(web_color,8)%>'>
<input type=hidden name=web_color_9 value='<%response.write web_var(web_color,9)%>'>
<input type=hidden name=web_color_10 value='<%response.write web_var(web_color,10)%>'>
<input type=hidden name=web_color_11 value='<%response.write web_var(web_color,11)%>'>
<input type=hidden name=web_color_12 value='<%response.write web_var(web_color,12)%>'>
<%
end if

if action="upload" then
  tt=1
%>
<tr>
<td>上传路径：</td>
<td><input type=text name=web_upload_1 value='<% response.write web_var(web_upload,1) %>' size=20 maxlength=20></td>
<td class=gray>&nbsp;为防有错，尽量不要修改此项</td>
</tr>
<tr>
<td>文件类型：</td>
<td><input type=text name=web_upload_2 value='<% response.write web_var(web_upload,2) %>' size=35 maxlength=50></td>
<td class=gray>&nbsp;多个类型用“,”分开</td>
</tr>
<tr>
<td>文件大小：</td>
<td><input type=text name=web_upload_3 value='<% response.write web_var(web_upload,3) %>' size=20 maxlength=10></td>
<td class=gray>&nbsp;单位为KB，最好小于500</td>
</tr>
<% else %>
<input type=hidden name=web_upload_1 value='<% response.write web_var(web_upload,1) %>'>
<input type=hidden name=web_upload_2 value='<% response.write web_var(web_upload,2) %>'>
<input type=hidden name=web_upload_3 value='<% response.write web_var(web_upload,3) %>'>
<%
end if

if action="info" then
  tt=1
%>
<tr><td colspan=3 class=red_3>&nbsp;过滤字符设置</td></tr>
<tr>
<td>非法字符：</td>
<td colspan=2><input type=text name=web_safety_1 value='<%response.write replace(web_var(web_safety,1),"'","")%>' size=30 maxlength=100>&nbsp;&nbsp;单引号(')和双引号(")已被系统过滤</td>
</tr>
<tr>
<td>密码允许：</td>
<td colspan=2><input type=text name=web_safety_2 value='<%response.write web_var(web_safety,2)%>' size=66 maxlength=100></td>
</tr>
<tr>
<td>注册禁用：</td>
<td colspan=2><input type=text name=web_safety_3 value='<%response.write web_var(web_safety,3)%>' size=66 maxlength=100></td>
</tr>
<tr>
<td>不健康字符：</td>
<td colspan=2><input type=text name=web_safety_4 value='<%response.write web_var(web_safety,4)%>' size=66 maxlength=100></td>
</tr>
<tr><td colspan=3 class=red_3>&nbsp;信息提示设置</td></tr>
<tr>
<td>外部提交：</td>
<td colspan=2><input type=text name=web_error_1 value='<%response.write web_var(web_error,1)%>' size=66 maxlength=100></td>
</tr>
<tr>
<td>未注册登陆：</td>
<td colspan=2><input type=text name=web_error_2 value='<%response.write web_var(web_error,2)%>' size=66 maxlength=100></td>
</tr>
<tr>
<td>支持信息：</td>
<td colspan=2><input type=text name=web_error_3 value='<%response.write web_var(web_error,3)%>' size=66 maxlength=100></td>
</tr>
<tr>
<td>网站底部：</td>
<td colspan=2><input type=text name=web_error_4 value='<%response.write web_var(web_error,4)%>' size=66 maxlength=100></td>
</tr>
<% else %>
<input type=hidden name=web_safety_1 value='<%response.write replace(web_var(web_safety,1),"'","")%>'>
<input type=hidden name=web_safety_2 value='<%response.write web_var(web_safety,2)%>'>
<input type=hidden name=web_safety_3 value='<%response.write web_var(web_safety,3)%>'>
<input type=hidden name=web_safety_4 value='<%response.write web_var(web_safety,4)%>'>
<input type=hidden name=web_error_1 value='<%response.write web_var(web_error,1)%>'>
<input type=hidden name=web_error_2 value='<%response.write web_var(web_error,2)%>'>
<input type=hidden name=web_error_3 value='<%response.write web_var(web_error,3)%>'>
<input type=hidden name=web_error_4 value='<%response.write web_var(web_error,4)%>'>
<%
end if

if action="down" then
  tt=1
%>
<tr>
<td>软件图片宽：</td>
<td><input type=text name=web_down_1 value='<%response.write web_var(web_down,1)%>' size=20 maxlength=10>&nbsp;像素</td>
<td class=gray>&nbsp;软件图片的宽度</td>
</tr>
<tr>
<td>软件图片高：</td>
<td><input type=text name=web_down_2 value='<%response.write web_var(web_down,2)%>' size=20 maxlength=10>&nbsp;像素</td>
<td class=gray>&nbsp;软件图片的高度</td>
</tr>
<tr>
<td>软件目录：</td>
<td><input type=text name=web_down_5 value='<%response.write web_var(web_down,5)%>' size=20 maxlength=20></td>
<td class=gray>&nbsp;</td>
</tr>
<tr>
<td>授权类型：</td>
<td colspan=2><input type=text name=web_down_3 value='<%response.write web_var(web_down,3)%>' size=66 maxlength=100></td>
</tr>
<tr>
<td>运行环境：</td>
<td colspan=2><input type=text name=web_down_4 value='<%response.write web_var(web_down,4)%>' size=66 maxlength=100></td>
</tr>
<% else %>
<input type=hidden name=web_down_1 value='<%response.write web_var(web_down,1)%>'>
<input type=hidden name=web_down_2 value='<%response.write web_var(web_down,2)%>'>
<input type=hidden name=web_down_5 value='<%response.write web_var(web_down,5)%>'>
<input type=hidden name=web_down_3 value='<%response.write web_var(web_down,3)%>'>
<input type=hidden name=web_down_4 value='<%response.write web_var(web_down,4)%>'>
<%
end if

if action="user" then
  tt=1
  tdim=split(user_power,"|")
  for i=0 to ubound(tdim)
%>
<tr>
<td>用户组 <%response.write i+1%>：</td>
<td><input type=text name=user_power_<%response.write i+1%> value='<%response.write tdim(i)%>' size=30 maxlength=20></td>
<td class=gray>&nbsp;</td>
</tr>
<%
  next
  erase tdim
%>
<input type=hidden name=user_power_num value='<%response.write i%>'>
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
<% else %>
<input type=hidden name=user_power value='<%response.write user_power%>'>
<%
end if

if action="forum" then
  tt=1
  udim=split(user_power,"|")
  tdim=split(forum_type,"|")
  for i=0 to ubound(tdim)
    t2=left(tdim(i),instr(tdim(i),":")-1)
%>
<tr>
<td>论坛分类 <%response.write i+1%>：</td>
<td><input type=text name=forum_type_<%response.write i+1%>_2 value='<%response.write right(tdim(i),len(tdim(i))-instr(tdim(i),":"))%>' size=20 maxlength=20></td>
<td class=gray>&nbsp;</td>
</tr>
<tr><td class=gray>论坛权限</td><td colspan=2><%
for j=0 to ubound(udim)
  response.write vbcrlf&"<input type=checkbox name=forum_type_"&i+1&"_1 value='"&j+1&"' class=bg_1"
  if instr(1,"."&t2&".","."&j+1&".")>0 then response.write " checked"
  response.write ">"&right(udim(j),len(udim(j))-instr(udim(j),":"))
next
%><input type=checkbox name=forum_type_<%response.write i+1%>_1 value='0' class=bg_1<%if instr(1,"."&t2&".",".0.")>0 then response.write " checked"%>>游客</td></tr><%
  next
  erase tdim
%>
<input type=hidden name=forum_type_num value='<%response.write i%>'>
<tr>
<td>新论坛分类：</td>
<td><input type=text name=forum_type_new_2 value='' size=30 maxlength=20></td>
<td class=gray>&nbsp;例：精华论坛</td>
</tr>
<tr><td class=gray>论坛权限</td><td colspan=2><%
for j=0 to ubound(udim)
  response.write vbcrlf&"<input type=checkbox name=forum_type_new_1 value='"&j+1&"' class=bg_1>"&right(udim(j),len(udim(j))-instr(udim(j),":"))
next
%><input type=checkbox name=forum_type_new_1 value='0' class=bg_1>游客</td></tr>
<tr>
<td colspan=3 class=htd><font class=red>论坛分类修改说明：</font>你可以添加、修改、删除论坛分类，但文本框中只能是以“文章学习”的形式（论坛分类名称）存在，否则将出错！！！
<font class=red>如用户组有删除或修改其权限后，请重新分配论坛权限！如论坛分类及权限有修改，请进入论坛管理以重新分配论坛类别！</font><br>
<font class=red_3>添加论坛分类</font>即在“新论坛分类”中按规则加入要新添的论坛分类，只能单个的增加；<br>
<font class=red_3>修改论坛分类</font>可修改所有的，排序对论坛的权限没有影响；<br>
<font class=red_3>删除论坛分类</font>可一次删除一个或多个，方法为将要删除的论坛分类框里的内容删除即可。</td>
</tr>
<%
  erase udim
else
%>
<input type=hidden name=forum_type value='<%response.write forum_type%>'>
<%
end if
%>
<!-- #include file="INCLUDE/config_forum.asp" -->
<% if not(isnumeric(forumid)) then call cookies_type("forum_id") %>
<!-- #include file="INCLUDE/jk_pagecute.asp" -->
<!-- #include file="include/conn.asp" -->
<%
' ====================
' Beyondest.Com v3.6.1
' http://beyondest.com
' ====================

call forum_first()

select case action
case "manage"
  tit=forumname&" [版面管理]"
case "isgood"
  tit=forumname&" [本版精华]"
case "move"
case else
  action=""
  tit=forumname
end select

call web_head(0,0,2,0,0)
'-----------------------------------center---------------------------------
dim rssum,nummer,thepages,viewpage,page,pageurl,page_cute_num,view_url,topic_head,del_temp,keyword,sea_type,sea_true,sea_write
dim forum_temp,id,username,icon,topic,tim,counter,re_counter,re_username,re_tim,istop,islock,isgood,folder_type,forumnid

if action="move" then call forum_moved(forumid,trim(request.querystring("view_id")))

if action="manage" then
  if format_user_power(login_username,login_mode,forumpower)<>"yes" then action=""
end if

pageurl="forum_list.asp?forum_id="&forumid&"&action="&action&"&"
rssum=0:thepages=0:viewpage=0:nummer=web_var(web_num,2):page_cute_num=web_var(web_num,3)
del_temp=0:forum_temp=""
keyword=code_form(request.querystring("keyword"))
sea_type=trim(request.querystring("sea_type"))
if (sea_type="topic" or sea_type="username" or sea_type="re_username") and len(keyword)>0 then
  sea_true="yes"
  sea_write=".搜索"
  pageurl=pageurl&"sea_type="&sea_type&"&keyword="&server.htmlencode(keyword)&"&"
else
  sea_true="no"
  sea_write=""
end if

select case action
case "manage"
  response.write forum_top("帖子列表 [版面管理"&sea_write&"]")
case "isgood"
  response.write forum_top("帖子列表 [精华列表"&sea_write&"]")
case else
  if sea_true="yes" then
    response.write forum_top("帖子列表 [搜索结果]")
  else
    response.write forum_top("帖子列表 （主题：<font class=red>"&forumtopicnum&"</font>）")
  end if
end select

%>
<script language=javascript>
<!--
function load_tree(f_id,v_id){
  var targetImg =eval("document.all.followImg" + v_id);
  var targetDiv =eval("document.all.follow" + v_id);
  if (targetImg.src.indexOf("nofollow")!=-1){return false;}
    if ("object"==typeof(targetImg)){
      if (targetDiv.style.display!='block'){
        targetDiv.style.display="block";
        targetImg.src="images/small/fk_minus.gif";
        if (targetImg.loaded=="no"){
          document.frames["hiddenframe"].location.replace("forum_loadtree.asp?forum_id="+f_id+"&view_id="+v_id);
        }
      }else{
      targetDiv.style.display="none";
      targetImg.src="images/small/fk_plus.gif";
    }
  }
}
-->
</script>
<iframe width=0 height=0 src='about:blank' id=hiddenframe></iframe>
<table border=0 width='98%'><tr><td align=left width='15%'><a href='forum_write.asp?forum_id=<%=forumid%>'><img src='images/<%=web_var(web_config,5)%>/new_topic.gif' align=absMiddle border=0 title='在 <%=forumname%> 里发表我的新贴'></a></td><td align=right width='85%'><table border=0><form action='?' method=get><input type=hidden name=forum_id value='<%=forumid%>'><input type=hidden name=action value='<%=action%>'><input type=hidden name=page value='<%=viewpage%>'><tr><td>论坛搜索：</td><td><select name=sea_type size=1><option value='topic'>按主题</option><option value='username'>按作者</option><option value='re_username'>按回复人</option></select></td><td><input type=text name=keyword size=20 maxlength=20></td><td>&nbsp;<input type=submit value='搜 索'></td></tr></table></td></tr></table>
<% response.write forum_table1 %>
<tr align=center<%response.write forum_table2%> height=25 >
<td rowap width='4%' class=end  background=images/<%=web_var(web_config,5)%>/bar_3_bg.gif>图</td>
<td rowap width='3%' class=end background=images/<%=web_var(web_config,5)%>/bar_3_bg.gif>例</td>
<td rowap width='48%' class=end background=images/<%=web_var(web_config,5)%>/bar_3_bg.gif>主题（每页&nbsp;<%response.write nummer%>&nbsp;贴&nbsp;&nbsp;点击&nbsp;<img src='IMAGES/SMALL/FK_PLUS.GIF' align=absMiddle border=0>&nbsp;可展开贴子列表）</td>
<td rowap width='12%' class=end background=images/<%=web_var(web_config,5)%>/bar_3_bg.gif>作者</td>
<td rowap width='4%' class=end background=images/<%=web_var(web_config,5)%>/bar_3_bg.gif>&nbsp;</td>
<td rowap width='7%' class=end background=images/<%=web_var(web_config,5)%>/bar_3_bg.gif>人气</td>
<td rowap width='22%' class=end background=images/<%=web_var(web_config,5)%>/bar_3_bg.gif>最后回复信息</td>
</tr>
<%
if action="manage" then
  response.write "<form name=del_form action='forum_isaction.asp?isaction=delete&forum_id="&forumid&"' method=post>"
end if

if action="isgood" then
  sql="select count(id) from bbs_topic where forum_id="&forumid&" and isgood=1"
  set rs=conn.execute(sql)
  forumtopicnum=rs(0)
  rs.close
end if
if sea_true="yes" then
  sql="select count(id) from bbs_topic where forum_id="&forumid&" and "&sea_type&" like '%"&keyword&"%'"
  set rs=conn.execute(sql)
  forumtopicnum=rs(0)
  rs.close
end if

sql="select * from bbs_topic where (forum_id="&forumid&" or istop=2) "
if action="isgood" then
  sql=sql&"and isgood=1"
end if
if sea_true="yes" then
  sql=sql&"and "&sea_type&" like '%"&keyword&"%'"
end if
sql=sql&" order by istop desc,re_tim desc,id desc"
set rs=conn.execute(sql)
if rs.eof and rs.bof then
  rssum=0
  response.write "<tr><td colspan=8 align=center height=50>本论坛暂时没有贴子。</td></tr>"
else
  rssum=forumtopicnum		'rs.recordcount
end if

if int(rssum)>0 then
  call format_pagecute()
  if int(viewpage)>1 then
    rs.move (viewpage-1)*nummer
  end if
  for i=1 to nummer
    if rs.eof then exit for
    folder_type="isok"
    forumnid=rs("forum_id")
    id=rs("id")
    username=rs("username")
    topic=rs("topic")
    icon=rs("icon")
    tim=rs("tim")
    counter=rs("counter")
    re_counter=rs("re_counter")
    re_username=rs("re_username")
    re_tim=rs("re_tim")
    istop=rs("istop")
    islock=rs("islock")
    isgood=rs("isgood")
    
    call forum_view()
    
    rs.movenext
  next
end if
rs.close:set rs=nothing

response.write "</table>"

if int(thepages)<1 then page_cute_num=1

response.write kong & forum_table1
%>
<tr height=25<%response.write forum_table3%>>
<td width='35%'>
主题：<font class=red><%response.write forumtopicnum%></font>&nbsp;
<%
if action<>"isgood" and sea_true<>"yes" then
  response.write "贴子总数：<font class=red>"&forumdatanum&"</font>&nbsp;"
end if
%>
页次：<font class=red><%response.write viewpage&"</font>/<font class=red>"&thepages%></font><td align=center>
分页：<% response.write jk_pagecute(nummer,thepages,viewpage,pageurl,page_cute_num,"#ff0000") %></td>
</td><td align=center width='25%'><% response.write forum_go() %></td>
</tr>
<% if action="manage" then %>
<script language=javascript src='STYLE/admin_del.js'></script>
<tr<%response.write format_table(3,1)%>><td height=25 align=center colspan=3>版面管理：　<input type=checkbox name=del_all value=1 onClick="selectall('<% response.write del_temp %>')" class=bg_1> 选中所有　<input type=submit value='删除所选' onclick="return suredel('<% response.write del_temp %>');"></td></tr>
</form>
<% end if %>
</table>
<table border=0 width='95%'>

<tr><td align=center colspan=2 height=30>
<%response.write web_var(web_config,1)%>论坛主题图例：&nbsp;
<%call is_type()%>
</td></tr>
</table>
<%
'---------------------------------center end-------------------------------
call web_end(0)
%>
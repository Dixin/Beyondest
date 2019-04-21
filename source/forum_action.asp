<!-- #include file="include/config_forum.asp" -->
<!-- #include file="include/jk_pagecute.asp" -->
<!-- #include file="include/conn.asp" -->
<%
'*******************************************************************
'
'                     Beyondest.Com v3.6.1
' 
'           http://beyondest.com
' 
'*******************************************************************

dim rssum,thepages,page,viewpage,sqladd,nummer,forum_temp,pageurl,usern
rssum=0:thepages=0:viewpage=1:nummer=web_var(web_num,1)
forum_temp="":pageurl=""

action=trim(request.querystring("action"))
select case action
case "hot"
  tit="论坛热贴"
  sql="select * from bbs_topic where re_counter>10 order by re_counter desc,id desc"
case "top"
  tit="论坛置顶"
  sql="select * from bbs_topic where is"&action&"<>0 order by istop desc,id desc"
case "good"
  tit="论坛精华"
  sql="select * from bbs_topic where is"&action&"=1 order by id desc"
case "tim"
  tit="回复新贴"
  sql="select top 100 * from bbs_topic order by re_tim desc,id desc"
case "user"
  usern=replace(trim(request.querystring("username")),"'","")
  if len(usern)<1 then
    call cookies_type("username")
  end if
  sql="select id from user_data where username='"&usern&"'"
  set rs=conn.execute(sql)
  if rs.eof and rs.bof then
    rs.close:set rs=nothing
    close_conn
    call cookies_type("username")
  end if
  rs.close
  tit="查看 "&usern&" 参与过的主题"
  pageurl="?action="&action&"&username="&usern&"&"
  sql="select bbs_topic.id,bbs_topic.forum_id,bbs_topic.username,bbs_topic.topic,bbs_topic.tim,bbs_topic.counter,bbs_topic.re_counter,bbs_topic.re_username,bbs_topic.re_tim,bbs_topic.istop,bbs_topic.islock,bbs_topic.isgood " & _
      "from bbs_data inner join bbs_topic on bbs_data.reply_id=bbs_topic.id where bbs_data.username='"&usern&"' group by bbs_topic.id,bbs_topic.forum_id,bbs_topic.username,bbs_topic.topic,bbs_topic.tim,bbs_topic.counter,bbs_topic.re_counter,bbs_topic.re_username,bbs_topic.re_tim,bbs_topic.istop,bbs_topic.islock,bbs_topic.isgood order by bbs_topic.id desc"
case "my"
  tit="我所参与过的主题"
  sql="select bbs_topic.id,bbs_topic.forum_id,bbs_topic.username,bbs_topic.topic,bbs_topic.tim,bbs_topic.counter,bbs_topic.re_counter,bbs_topic.re_username,bbs_topic.re_tim,bbs_topic.istop,bbs_topic.islock,bbs_topic.isgood " & _
      "from bbs_data inner join bbs_topic on bbs_data.reply_id=bbs_topic.id where bbs_data.username='"&login_username&"' group by bbs_topic.id,bbs_topic.forum_id,bbs_topic.username,bbs_topic.topic,bbs_topic.tim,bbs_topic.counter,bbs_topic.re_counter,bbs_topic.re_username,bbs_topic.re_tim,bbs_topic.istop,bbs_topic.islock,bbs_topic.isgood order by bbs_topic.id desc"
case else
  tit="论坛新贴"
  sql="select top 100 * from bbs_topic order by id desc"
end select
if pageurl="" then pageurl="?action="&action&"&"

call web_head(0,0,0,0,0)
'------------------------------------left----------------------------------
call format_login()
response.write left_action("jt13",4)
'----------------------------------left end--------------------------------
call web_center(0)
'-----------------------------------center---------------------------------
response.write ukong

set rs=server.createobject("adodb.recordset")
rs.open sql,conn,1,1

if not(rs.eof and rs.bof) then
  rssum=rs.recordcount
end if

call format_pagecute()
if int(viewpage)>1 then
  rs.move (viewpage-1)*nummer
end if
for i=1 to nummer
  if rs.eof then exit for
  forum_temp=forum_temp&forum_view()
  rs.movenext
next
rs.close:set rs=nothing

response.write forum_table1
%>
<tr height=30 bgcolor=<%=web_var(web_color,6)%> align=center>
<td width='75%'><font class=red_3><b><%response.write tit%></b></font>&nbsp;&nbsp;&nbsp;
共&nbsp;<font class=red><%response.write rssum%></font>&nbsp;贴&nbsp;┋&nbsp;
每&nbsp;<font class=red><%response.write nummer%></font>&nbsp;页&nbsp;┋&nbsp;
共&nbsp;<font class=red><%response.write thepages%></font>&nbsp;页&nbsp;┋&nbsp;
这是第&nbsp;<font class=red><%response.write viewpage%></font>&nbsp;页</td>
</tr>
</table>
<% response.write kong & forum_table1 %>
<tr align=center<%response.write forum_table2%> height=25>
<td width='5%' background=images/<%=web_var(web_config,5)%>/bar_3_bg.gif>&nbsp;</td>
<td width='58%' background=images/<%=web_var(web_config,5)%>/bar_3_bg.gif><font class=end>论坛主题</font></td>
<td width='14%' background=images/<%=web_var(web_config,5)%>/bar_3_bg.gif><font class=end>作者</font></td>
<td width='9%' background=images/<%=web_var(web_config,5)%>/bar_3_bg.gif><font class=end>人气</font></td>
<td width='14%' background=images/<%=web_var(web_config,5)%>/bar_3_bg.gif><font class=end>最后回复</font></td>
</tr>
<% response.write forum_temp %>
</table>
<br>
<% response.write forum_table1 %>
<tr height=30 bgcolor=<%=web_var(web_color,6)%>>
<td width='70%'>&nbsp;分页：<%response.write jk_pagecute(nummer,thepages,viewpage,pageurl,5,"#ff0000")%></td>
<td width='30%' align=center><% response.write forum_go() %></td>
</tr>
<tr<%response.write forum_table4%>><td align=center height=30 colspan=2>
<%response.write img_small("isok")%>&nbsp;开放的主题&nbsp;&nbsp;
<%response.write img_small("ishot")%>&nbsp;回复超过10贴&nbsp;&nbsp;
<%response.write img_small("islock")%>&nbsp;锁定的主题&nbsp;&nbsp;
<%response.write img_small("istop")%>&nbsp;固定顶端的主题&nbsp;&nbsp;
<%response.write img_small("isgood")%>&nbsp;精华帖子
</td></tr>
</table>
<br>
<%
'---------------------------------center end-------------------------------
call web_end(0)

function forum_view()
  dim view_url,topic_head,forumid,id,username,topic,tim,counter,re_counter,re_username,re_tim,istop,islock,isgood,folder_type,reply_count
  folder_type="isok"
  id=rs("id")
  username=rs("username")
  topic=rs("topic")
  tim=rs("tim")
  counter=rs("counter")
  re_counter=rs("re_counter")
  re_username=rs("re_username")
  re_tim=rs("re_tim")
  istop=rs("istop")
  islock=rs("islock")
  isgood=rs("isgood")
  
  select case int(istop)
  case 1
    folder_type="istop"
  case 2
    folder_type="istops"
  case else
    if int(isgood)=1 then
      folder_type="isgood"
    else
      if int(islock)=1 then
        folder_type="islock"
      elseif int(re_counter)>=10 then
        folder_type="ishot"
      end if
    end if
  end select
  
  forumid=rs("forum_id")
  view_url="forum_view.asp?forum_id="&forumid&"&view_id="&id
  if int(re_counter)>0 then
    topic_head="<img loaded=no src='images/small/fk_plus.gif' border=0>"
  else
    topic_head="<img src='images/small/fk_minus.gif' border=0>"
  end if
  forum_view=vbcrlf & "<tr align=center"&forum_table4&"><td><img src='images/small/"&folder_type&".gif' border=0></td>" & _
	     vbcrlf & "<td align=left>"&topic_head&"<a href='"&view_url&"' title='主题："&code_html(topic,1,0)&"<br>发贴时间："&tim&"<br>最后回复："&re_username&"<br>回复时间："&re_tim&"'>"&code_html(topic,1,25)&"</a>&nbsp;"&index_pagecute(view_url,re_counter+1,web_var(web_num,3),"#cc3300")&"</td>" & _
	     vbcrlf & "<td>"&format_user_view(username,1,"")&"</td>" & _
	     vbcrlf & "<td class=timtd>"&re_counter&"/"&counter&"</td>" & _
	     vbcrlf & "<td>"&format_user_view(re_username,1,"")&"</td></tr>"
end function
%>
<!-- #include file="config.asp" -->
<!-- #include file="skin.asp" -->
<%
'*******************************************************************
'
'                     Beyondest.Com V3.6 Demo版
' 
'           网址：http://www.beyondest.com
' 
'*******************************************************************

index_url="main"
tit_fir=format_menu(index_url)

dim us,table1,table2,table3,table4
us="fk2"
table1=format_table(1,3)
table2=format_table(3,2)
table3=format_table(3,1)
table4=format_table(3,5)

sub user_data_top(utt,ijt,sh,n_num)
  '积分排行	integral
  '发贴排行	bbs_counter
  '财富排行	emoney
  dim temp1
  temp1=vbcrlf&"<table border=0 width='99%' cellspacing=2 cellpadding=0 align=center class=tf><tr height=0><td width='75%'></td><td width='25%'></td></tr>"
  sql="select top "&n_num&" username,"&utt&" from user_data order by "&utt&" desc,id"
  set rs=conn.execute(sql)
  do while not rs.eof
    temp1=temp1&vbcrlf&"<tr><td height="&space_mod&" class=bw>"&img_small(ijt)&format_user_view(rs("username"),1,"")&"</td><td align=center class=red_3>"&rs(utt)&"</td></tr>"
    rs.movenext
  loop
  rs.close
  temp1=temp1&"</table>"
  call left_btype(temp1,utt,sh,12)
end sub

sub vote_type(t1,t2,t3,t4)
  call left_btype("<script language=javascript src='vote.asp?id="&t1&"&types="&t2&"&mcolor="&t3&"&bgcolor="&t4&"'></script>","vote",1,13)
end sub

sub left_type(ltv,pn,ltt)
  if ltt=1 then response.write "<table border=0 width='100%' cellspacing=0 cellpadding=0 align=center><tr><td align=center>"
  response.write kong&format_barc("<img src='images/"&web_var(web_config,5)&"/left_"&pn&".gif' border=0>",ltv,2,0,10)
  if ltt=1 then response.write "</td></tr></table>"
end sub

sub left_btype(ltv,pn,ltt,lbicon)
  if ltt=1 then response.write "<table border=0 width='100%' cellspacing=0 cellpadding=0 align=center><tr><td align=center>"
  response.write kong&format_barc("<img src='images/"&web_var(web_config,5)&"/left_"&pn&".gif' border=0>",ltv,2,0,lbicon)
  if ltt=1 then response.write "</td></tr></table>"
end sub

sub main_down_pic(nnhead,dsql,nt,n_num,c_num)
  dim temp1,nid,name,nhead:nhead=nnhead
  temp1="<table border=0 width='100%' cellspacing=0 cellpadding=2><tr align=center valign=top>"
  sql="select top "&n_num&" id,name,tim,pic from down where hidden=1"&dsql
  select case nt
  case "hot"
    sql=sql&" order by counter desc,id desc"
    if nhead="" then nhead="热点排行"
  case "good"
    sql=sql&" and types=5 order by id desc"
    if nhead="" then nhead="精彩推荐"
  case else
    sql=sql&" order by id desc"
    if nhead="" then nhead="最新上传"
  end select
  set rs=conn.execute(sql)
  do while not rs.eof
    name=rs("name"):nid=rs("id")
    temp1=temp1&vbcrlf&"<td width='"&int(100\n_num)&"%'><table border=0 cellspacing=0 cellpadding=2 width='100%' class=tf><tr><td align=center><a href='down_view.asp?id="&nid&"'"&atb&"><img src='"&web_var(web_upload,1)&rs("pic")&"' border=0 width="&web_var(web_down,1)&" height="&web_var(web_down,2)&"></a></td></tr>" & _
	  vbcrlf&"<tr><td align=center class=bw><a href='down_view.asp?id="&nid&"'"&atb&" class=red_3><b>"&code_html(name,1,0)&"</b></a></td></tr></table></td>"
    rs.movenext
  loop
  temp1=temp1&"</tr></table>"
  response.write kong&format_bar("<font class="&sk_class&"><b>"&nhead&"</b></font>",temp1,sk_bar,0,0,"||","")
end sub

sub main_down(n_jt,nnhead,nmore,nsql,nt,n_num,n_m,c_num,et,tt)
  dim temp1,nid,name,tim,counter,nhead:nhead=nnhead
  if n_jt<>"" then n_jt=img_small(n_jt)
  temp1=vbcrlf&"<table border=0 width='100%' cellspacing=0 cellpadding=2 class=tf>"
  sql="select top "&n_num+n_m&" id,name,username,tim,counter from down where hidden=1"&nsql
  select case nt
  case "hot"
    sql=sql&" order by counter desc,id desc"
    if nhead="" then nhead="下载排行"
  case "good"
    sql=sql&" and types=5 order by id desc"
    if nhead="" then nhead="精彩推荐"
  case else
    sql=sql&" order by id desc"
    if nhead="" then nhead="音乐更新"
  end select
  set rs=conn.execute(sql)
  if n_m>0 then rs.move(n_m)
  do while not rs.eof
    name=rs("name"):tim=rs("tim"):counter=rs("counter")
    temp1=temp1&vbcrlf&"<tr><td height="&space_mod&" class=bw>"&n_jt&"<a href='down_view.asp?id="&rs("id")&"'"&atb&" title='音乐名称："&code_html(name,1,0)&"<br>发 布 人："&rs("username")&"<br>下载人次："&counter&"<br>整理时间："&time_type(tim,88)&"'>"&code_html(name,1,c_num)&"</a>"
    if tt>0 then temp1=temp1&format_end(et,time_type(tim,tt)&"，点击：<font class=blue>"&counter&"</font>")
    temp1=temp1&"</td></tr>"
    rs.movenext
  loop
  rs.close
  temp1=temp1&vbcrlf&"</table>"
  response.write format_barc("<a href=down.asp><font class="&sk_class&"><b>"&nhead&"</b></font></a>",temp1&kong,3,0,5)
end sub

sub main_news(n_jt,n_num,c_num,et,timt)
  dim temp1,topic,oid,ooid,pic
  if n_jt<>"" then n_jt=img_small(n_jt)
  temp1="<table border=0 width='100%' cellspacing=0 cellpadding=2 class=tf>"
  sql="select top "&n_num&" id,topic,tim,counter from news where hidden=1 order by id desc"
  set rs=conn.execute(sql)
  do while not rs.eof
    topic=rs("topic")
    temp1=temp1&vbcrlf&"<tr><td height="&space_mod&">"&n_jt&"<a href='news_view.asp?id="&rs("id")&"'"&atb&" title='"&code_html(topic,1,0)&"'>"&code_html(topic,1,c_num)&"</a>"
    if et>-1 then temp1=temp1&format_end(et,"<font class=gray>"&time_type(rs("tim"),timt)&"，点击：</font><font class=blue>"&rs("counter")&"</font>")
    temp1=temp1&"</td></tr>"
    rs.movenext
  loop
  rs.close
sql="select top 1 id,topic,pic from news where hidden=1 and istop=1 and ispic=1 order by id desc"
  set rs=conn.execute(sql)
  if not(rs.eof) then
    ooid=rs("id")
    if et=1 then oid=ooid
    pic="<table border=0><tr><td align=center valign=middle><a href='news_view.asp?id="&ooid&"' target='_blank'><img src='"&url_true(web_var(web_down,5),rs("pic"))&"' border='0' width="&web_var(web_num,8)&" height="&web_var(web_num,7)&"></a></td></tr><tr><td align=center><a href='news_view.asp?id="&ooid&"' target='_blank'>"&code_html(rs("topic"),1,c_num)&"</a></td></tr></table>"
  end if
  rs.close
  temp1="<table width='100%'  border='0' cellspacing='0' cellpadding='0'><tr><td width='40%' align=center>"&pic&"</td><td width=1 bgcolor="&web_var(web_color,3)&"></td><td>"&temp1&"</table></td></tr></table>"
  
  response.write format_barc("<a href='news.asp'><b><font class=end>最近新闻</font></b></a>",temp1&kong,3,0,7)
end sub

sub main_article(n_jt,n_num,c_num,et,timt)
  dim temp1,topic
  if n_jt<>"" then n_jt=img_small(n_jt)
  temp1="<table border=0 width='100%' cellspacing=0 cellpadding=2 class=tf>"
  sql="select top "&n_num&" id,topic,tim,counter from article where hidden=1 order by id desc"
  set rs=conn.execute(sql)
  do while not rs.eof
    topic=rs("topic")
    temp1=temp1&vbcrlf&"<tr><td height="&space_mod&">"&n_jt&"<a href='article_view.asp?id="&rs("id")&"'"&atb&" title='"&code_html(topic,1,0)&"'>"&code_html(topic,1,c_num)&"</a>"
    if et>-1 then temp1=temp1&format_end(et,"<font class=gray>"&time_type(rs("tim"),timt)&"，点击：</font><font class=blue>"&rs("counter")&"</font>")
    temp1=temp1&"</td></tr>"
    rs.movenext
  loop
  rs.close
  temp1=temp1&"</table>"
  response.write format_barc("<a href='article.asp'><b><font class=end>最新资料</font></b></a>",temp1&kong,3,0,2)
end sub

sub main_video(v_jt,v_num,vv_num,vet,vtimt)
  dim temp1,topic
  if v_jt<>"" then v_jt=img_small(v_jt)
  temp1="<table border=0 width='100%' cellspacing=0 cellpadding=2 class=tf>"
  sql="select top "&v_num&" * from gallery where types='film' and hidden=1 order by id desc"
  set rs=conn.execute(sql)
  do while not rs.eof
    topic=rs("name")
    temp1=temp1&vbcrlf&"<tr><td height="&space_mod&">"&v_jt&"<a href='gallery.asp?types=view&action=film&id="&rs("id")&"'"&atb&" title='"&code_html(topic,1,0)&"'>"&code_html(topic,1,vv_num)&"</a>"
    if vet>-1 then temp1=temp1&format_end(vet,"<font class=gray>"&time_type(rs("tim"),vtimt)&"，点击：</font><font class=blue>"&rs("counter")&"</font>")
    temp1=temp1&"</td></tr>"
    rs.movenext
  loop
  rs.close
  temp1=temp1&"</table>"
  response.write format_barc("<a href='vouch.asp'><b><font class=end>最新视频</font></b></a>",temp1&kong,3,0,4)
end sub

sub main_forum(n_jt,n_num,c_num,et,timt)
  dim temp1,topic
  if n_jt<>"" then n_jt=img_small(n_jt)
  temp1="<table border=0 width='100%' cellspacing=0 cellpadding=2 class=tf>"
  sql="select top "&n_num&" forum_id,id,topic,tim,counter,re_counter from bbs_topic order by id desc"
  set rs=conn.execute(sql)
  do while not rs.eof
    topic=rs("topic")
    temp1=temp1&vbcrlf&"<tr><td height="&space_mod&">"&n_jt&"<a href='forum_view.asp?forum_id="&rs("forum_id")&"&view_id="&rs("id")&"'"&atb&" title='"&code_html(topic,1,0)&"'>"&code_html(topic,1,c_num)&"</a>"&format_end(et,"<font class=gray>"&time_type(rs("tim"),timt)&"，</font><font class=blue>"&rs("re_counter")&"</font>|<font class=blue>"&rs("counter")&"</font>")&"</td></tr>"
    rs.movenext
  loop
  rs.close
  temp1=temp1&"</table>"
  response.write format_barc("<a href='forum.asp'><b><font class=end>论坛新贴</font></b></a>",temp1&kong,3,0,3)
end sub


sub main_update_view()
  dim temp1,topic
  temp1="<table border=0 width='100%' cellspacing=0 cellpadding=2 class=tf><tr><td></td><td  width='90%'>"&ukong&"</td><td></td></tr>"
  sql="select top 1 * from bbs_cast where sort='news' order by id desc"
  set rs=conn.execute(sql)
  do while not rs.eof
    topic=rs("topic")
    temp1=temp1&vbcrlf&"<tr><td></td><td align=center background='images/"&web_var(web_config,5)&"/roll_bg.gif'><marquee scrollamount=2 direction=up width='100%' height='120' onMouseOver=this.stop() onMouseOut=this.start()><center><a href='update.asp?action=news&id="&rs("id")&"'>"&code_html(topic,1,0)&"</center></a>"&code_jk(rs("word"))&"</marquee></td><td></td></tr>"
    rs.movenext
  loop
  rs.close
  temp1=temp1&"</table>"
  response.write format_barc("<a href='update.asp'><b><font class=end>网站更新</font></b></a>",temp1,2,0,1)
end sub

sub main_update(n_jt,n_num,c_num,et,timt)
  dim temp1,topic
  if n_jt<>"" then n_jt=img_small(n_jt)
  temp1="<table border=0 width='100%' cellspacing=0 cellpadding=2>"
  sql="select top "&n_num&" id,topic,tim from bbs_cast where sort='news' order by id desc"
  set rs=conn.execute(sql)
  do while not rs.eof
    topic=rs("topic")
    temp1=temp1&vbcrlf&"<tr align=center><td height="&space_mod&">"&n_jt&"<a href='update.asp?action=news&id="&rs("id")&"' title='"&code_html(topic,1,0)&"'>"&code_html(topic,1,c_num)&"</a>"&format_end(et,"<font class=gray>"&time_type(rs("tim"),timt)&"</font>")&"</td></tr>"
    rs.movenext
  loop
  rs.close
  temp1=temp1&"</table>"
  response.write kong&format_bar("<a href='update.asp'><b><font class=end>网站更新</font></b></a>",temp1,sk_bar,0,0,"|"&web_var(web_color,1)&"|80","")&"</td>"
end sub



sub main_shop()
  dim sql,rs
%><table border=0 width='98%' cellspacing=0 cellpadding=0><tr valign=top><%
  sql="select top 3 id,name,serial,brand,stock,smallimg,price_1,price_2,remark_1 from product where isgood=1 and hidden=1 order by tim desc,id desc"
  set rs=conn.execute(sql)
  do while not rs.eof
    response.write vbcrlf&"<td width='33%' align=center>"&shop_view()&"</td>"
    rs.movenext
  loop
  rs.close:set rs=nothing
%></tr></table><%
end sub

sub links_main(lt,nummer)
  dim temp1,nname
  temp1="<table border=0 width='100%' cellspacing=0 cellpadding=0>"
  sql="select * from links where sort='"&lt&"' and hidden=1 order by orders"
  set rs=conn.execute(sql)
  do while not rs.eof
    temp1=temp1&vbcrlf&"<tr align=center height=40>"
    for i=1 to nummer
      if rs.eof then exit for
      temp1=temp1&vbcrlf&"<td width='16%'><a href='"&rs("url")&"' target=_blank>"
      if lt="txt" then
        temp1=temp1&code_html(rs("nname"),1,0)
      else
        temp1=temp1&"<img src='"&rs("pic")&"' border=0 width=88 height=31 title='"&code_html(rs("nname"),1,0)&"'>"
      end if
      temp1=temp1&"</a></td>"
      rs.movenext
    next
    temp1=temp1&vbcrlf&"</tr>"
  loop
  rs.close
  temp1=temp1&"</table>"
  response.write temp1
end sub

sub links_maini(lt,nummer)
  dim temp1,nname
  temp1="<table border=0 width='100%' cellspacing=0 cellpadding=0>"
  sql="select * from links where sort='"&lt&"' and hidden=1 order by orders"
  set rs=conn.execute(sql)
  do while not rs.eof
    temp1=temp1&vbcrlf&"<tr align=center height=40>"
    for i=1 to nummer
      if rs.eof then exit for
      temp1=temp1&vbcrlf&"<td width='16%'><a href='"&rs("url")&"' target=_blank>"
      if lt="txt" then
        temp1=temp1&code_html(rs("nname"),1,0)
      else
        temp1=temp1&"<img src='"&rs("pic")&"' border=0 width=88 height=31 title='"&code_html(rs("nname"),1,0)&"'>"
      end if
      temp1=temp1&"</a></td>"
      rs.movenext
    next
    temp1=temp1&vbcrlf&"</tr>"
  loop
  rs.close
  temp1=temp1&"</table>"
  response.write format_barc("<a href=links.asp><font class=end><b>合作站点</b></font></a>",temp1,1,1,4)
end sub

sub web_search(wt)
  dim temp11
  if wt=1 then response.write ukong
  temp11="<table border=0><form action='search.asp' method=get><tr><td>"&img_small("new")&"关键词：</td><td>&nbsp;<input type=text name=keyword value='' size=20 maxlength=20></td><td>&nbsp;&nbsp;<select name=sea_type><option value='down'>音乐</option><option value='forum'>论坛</option><option value='news'>新闻</option><option value='article'>资料</option><option value='paste'>壁纸</option><option value='flash'>Flash</option><option value='website'>网站</option></select>&nbsp;&nbsp;</td><td><input type=checkbox name=celerity value='yes' ></td><td>快速搜索&nbsp;&nbsp;</td><td valign=top><input type=image src='images/small/web_sea.gif' border=0></td><td>&nbsp;&nbsp;<a href='search.asp?action=help' title='多功能搜索'>搜索帮助</a></td></tr></form></table>"
  response.write format_barc("<font class=end><b>站内搜索</b></font>",temp11,1,1,3)

end sub

sub news_fpic(d_num,t_num,w_num,nft)
  dim temp1,topic,pic,tt,rs,sql,wnum,ispic,ooid
  wnum=w_num
  sql="select top 1 id,topic,pic from news where hidden=1 and istop=1 and ispic=1 order by id desc"
  set rs=conn.execute(sql)
  if not(rs.eof) then
    ooid=rs("id")
    if nft=1 then oid=ooid
    pic="<table border=0><tr><td align=center><a href='news_view.asp?id="&ooid&"'><img src='"&url_true(web_var(web_upload,1),rs("pic"))&"' border='0' width="&web_var(web_num,7)&" height="&web_var(web_num,8)&"></a></td></tr><tr><td align=center><a href='news_view.asp?id="&ooid&"'>"&code_html(rs("topic"),1,d_num)&"</a></td></tr></table>"
  end if
  rs.close
  sql="select top "&t_num&" id,topic,pic,ispic from news where hidden=1 and istop=1"
  if ooid<>0 then sql=sql&" and id<>"&ooid
  sql=sql&" order by id desc"
  set rs=conn.execute(sql)
  do while not rs.eof
    topic=rs("topic")
    if rs("ispic")=true then wnum=wnum-2:ispic=sk_img
    temp1=temp1&"<tr><td height="&space_mod&">"&img_small("jt0")&"<a href='news_view.asp?id="&rs("id")&"'>"&code_html(topic,1,wnum)&"</a>"&ispic&"</td></tr>"
    rs.movenext
  loop
  rs.close:set rs=nothing
  temp1="<table border=0 width='100%'><tr>" & _
        "<td width='30%' align=center>"&pic&"</td>" & _
        "<td width='70%'><table border=0>"&temp1&"</table></td>" & _
        "</tr></table>"
 response.write kong&format_bar("<font class="&sk_class&"><b>最新图片新闻</b></font>",temp1,sk_bar,0,0,"||","")
end sub


sub main_best()
  dim temp1,nid,n_jt,name,tim,counter,nhead
  n_jt=img_small("jt0")
  temp1=vbcrlf&"<table border=0 width='100%' cellspacing=0 cellpadding=0><tr valign=top>"


  sql="select top 1 id,name,counter,spic,tim,username,remark from gallery where hidden=1 and istop=1 and types='film' order by counter desc,id desc"
  nhead="精彩推荐"
  set rs=conn.execute(sql)
  do while not rs.eof
    temp1=temp1&vbcrlf&"<td class=bw width='33%'>"&kong&"&nbsp;&nbsp;推荐视频：<br>&nbsp;&nbsp;<a href='gallery.asp?types=view&action=film&id="&rs("id")&"' title=''  target='_blank'><img src='"&url_true("images/video",rs("spic"))&"' border='0' width=150 height=150></a>"
    temp1=temp1&"<br>&nbsp;&nbsp;"&n_jt&"视频名称："&code_html(rs("name"),1,0)&"<br>&nbsp;&nbsp;"&n_jt&"发 布 人："&rs("username")&"<br>&nbsp;&nbsp;"&n_jt&"点击次数："&rs("counter")&"<br>&nbsp;&nbsp;"&img_small("jt12")&left(code_jk(rs("remark")),12)&"……"
    temp1=temp1&"</td>"
    rs.movenext
  loop
  rs.close
  temp1=temp1&"<td width=1 bgcolor="&web_var(web_color,3)&"></td>"

  sql="select top 1 s_id,c_id,s_name,pic,intro from jk_sort where istop=1 order by s_id desc"
set rs=conn.execute(sql)
  do while not rs.eof
    temp1=temp1&"<td align=left class=bw  >"&kong&"&nbsp;&nbsp;推荐专辑：<br>&nbsp;&nbsp;<a href='down_list.asp?c_id="&rs("c_id")&"&s_id="&rs("s_id")&"' target='_blank'><img src=images/down/"&rs("pic")&".jpg width=150 height=150 border=0></a>"
    temp1=temp1&"<br>&nbsp;&nbsp;"&n_jt&rs("s_name")&"<table width=180 border=0  cellspacing=0 cellpadding=0><tr><td>&nbsp;&nbsp;</td><td>"&img_small("jt12")&left(rs("intro"),45)&"……</td></tr></table>"
    temp1=temp1&"</td>"
    rs.movenext
  loop
  rs.close
  temp1=temp1&"<td width=1 bgcolor="&web_var(web_color,3)&"></td>"

  sql="select top 1 * from down where hidden=1 and types=5 order by id desc"
set rs=conn.execute(sql)
  do while not rs.eof
    temp1=temp1&"<td align=left class=bw  width='33%'>"&kong&"&nbsp;&nbsp;推荐音乐：<br>&nbsp;&nbsp;<a href='down_view.asp?id="&rs("id")&"'  target='_blank'><img src=images/down/"&rs("pic")&" width=150 height=150 border=0></a>"
    temp1=temp1&"<br>&nbsp;&nbsp;"&n_jt&"音乐名称："&code_html(rs("name"),1,0)&"<br>&nbsp;&nbsp;"&n_jt&"发 布 人："&rs("username")&"<br>&nbsp;&nbsp;"&n_jt&"点击次数："&rs("counter")&"<br>&nbsp;&nbsp;"&img_small("jt12")&left(code_jk(rs("remark")),12)&"……"
    temp1=temp1&"</td>"
    rs.movenext
  loop
  rs.close



  temp1=temp1&vbcrlf&"</tr></table>"
  response.write format_barc("<a href=down.asp><font class="&sk_class&"><b>"&nhead&"</b></font></a>",temp1&kong,3,0,5)
end sub
%>
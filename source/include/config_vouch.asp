<!-- #include file="config.asp" -->
<!-- #include file="config_nsort.asp" -->
<!-- #include file="skin.asp" -->
<%
' ====================
' Beyondest.Com v3.6.1
' http://beyondest.com
' ====================

dim pageurl,nummer,sqladd,page,rssum,thepages,viewpage,keyword,sea_type,upload_file,id,pic_width,pic_height
pic_width=web_var(web_num,7):pic_height=web_var(web_num,8):upload_file=web_var(web_down,5)
rssum=0:nummer=web_var(web_num,2):thepages=0:viewpage=1
sk_bar=11
sk_class="end"
id=trim(request.querystring("id"))
if not(isnumeric(id)) then id=0
index_url="vouch"
tit_fir=""
call cid_sid()

sub vouch_left(vjt,vjt2)
  dim temp1
  upload_file=web_var(web_down,6)
  if vjt<>"" then vjt=img_small(vjt)
  if vjt2<>"" then vjt2=img_small(vjt2)
  temp1=vbcrlf&"<table border=0 width='80%' cellspacing=0 cellpadding=2 align=center>" & _
	vbcrlf&"<tr height=5><td></td></tr>" & _
	vbcrlf&"<tr><td>"&vjt&"<a href='gallery.asp?action=paste'>桌面壁纸</a></td></tr>" & _
	vbcrlf&"<tr><td>"&vjt&"<a href='gallery.asp?action=flash'>Flash MTV</a></td></tr>" & _
	vbcrlf&"<tr><td>"&vjt&"<a href='gallery.asp?action=film'>精彩视频</a></td></tr>" & _
	vbcrlf&"<tr><td>"&vjt&"<a href='gallery.asp?action=baner'>Beyond相册</a></td></tr>"& _
	vbcrlf&"<tr><td>"&vjt&"<a href='website.asp'>精彩网站</a></td></tr>"& _
        vbcrlf&"<tr><td align=right>"&vjt2&"<a href='user_put.asp?action=gallery'>我要上传图片</a></td></tr>" & _
        vbcrlf&"<tr><td align=right>"&vjt2&"<a href='user_put.asp?action=website'>我要推荐网站</a></td></tr>" & _
	vbcrlf&"</table>"
  call vouch_skin("精彩栏目特别推荐",temp1,"",1)
end sub

sub vouch_skin(t1,t2,t3,t4)
  if t4=1 then response.write "<table border=0 width='100%' cellspacing=0 cellpadding=0 align=center><tr><td align=center>"
  response.write format_barc("<font class="&sk_class&"><b>"&t1&"</b></font>",t2,2,0,8)
  if t4=1 then response.write "</td></tr></table>"
end sub

sub web_site_type()
%>
<table border=0>
<tr height=5><td width='3%'></td><td width='77%'></td><td width='30%'></td></tr>
<tr>
<td colspan=2><%response.write img_small("jt0")%><b><a href='?c_id=<%response.write cid%>&s_id=<%response.write sid%>&action=view&id=<%response.write nid%>' target=_blank><%response.write name%></a></b></td>
<td rowspan=2 align=center><img src='images/<%=rs("pic") %>' border=0></td>
</tr>
<tr><td></td><td>
  <table border=0 width='98%'>
  <tr><td width='18%'></td><td width='30%'</td><td width='18%'></td><td width='34%'></td></tr>
  <tr><td>国家地区：</td><td><%response.write rs("country")%></td><td>站点语言：</td><td><%response.write rs("lang")%></td></tr>
  <tr><td>推 荐 人：</td><td><%response.write format_user_view(rs("username"),1,"")%></td><td>添加时间：</td><td><%response.write time_type(rs("tim"),8)%></td></tr>
  <tr><td>站点属性：</td><td><%
if rs("isgood")=true then
  response.write "<font class=red_3>推荐</font>"
else
  response.write "普通"
end if
  %></td><td>浏览人气：</td><td class=red><%response.write rs("counter")%></td></tr>
  <tr><td>网站地址：</td><td colspan=3><a href='?c_id=<%response.write cid%>&s_id=<%response.write sid%>&action=view&id=<%response.write nid%>' target=_blank><%response.write url%></a></td></tr>
  <tr><td colspan=2>网站介绍：</td><td colspan=2 align=center><a href="javascript:window.external.AddFavorite('<%response.write url%>','<%response.write name%>')" style='target: ' _self?>〖加入收藏夹〗</a></td></tr>
  <tr><td colspan=4><table border=0 width='94%' align=center><tr><td><%response.write code_html(rs("remark"),1,0)%></td></tr></table></td></tr>
  </table>
</td></tr></table>
<%
end sub

sub gallery_new(gtype,n_num)
  dim rs,sql,pic_temp,pic,name
  response.write "<table border=0>"
  sql="select top "&n_num&" id,name,pic,c_id,s_id from gallery where types='"&gtype&"' order by id desc"
  set rs=conn.execute(sql)
  do while not rs.eof
    pic=rs("pic"):name=rs("name")
    select case gtype
    case "paste"
      pic_temp="<table border=0><tr><td align=center><a href='?action=view&c_id="&rs("c_id")&"&s_id="&rs("s_id")&"&id="&rs("id")&"'><img src='"&upload_file&pic&"' border=0 width="&pic_width&" height="&pic_height&"></a></td></tr><tr><td align=center>"&code_html(name,1,10)&"</td></tr></table>"
      response.write "<td width="&pic_width+30&" height="&pic_height+50&">"&format_k(pic_temp,1,5,pic_width+10,pic_height+30)&"</td>"
    case "logo"
      response.write "<td width=98><table border=0><tr><td align=center><img src='"&upload_file&pic&"' border=0 width=88 height=31></td></tr><tr><td align=center><font title='"&name&"'>"&code_html(name,1,10)&"</font></td></tr></table></td>"
    end select
    rs.movenext
  loop
  rs.close:set rs=nothing
  response.write "</table>"
end sub

sub film_view(ispic,width,height)
  dim file_type:file_type=right(ispic,4)
  if instr(1,file_type,".")>0 then file_type=right(file_type,len(file_type)-instr(1,file_type,"."))
  file_type=lcase(file_type)
  if lcase(left(ispic,3))="mms" then file_type="mms"
  select case file_type
  case "rm","ram","rmvb","ra"
%>
<object id="video2" classid="clsid:CFCDAA03-8BE4-11cf-B84B-0020AFBBCCFA" width=530 height=350>
  <param name="_ExtentX" value="11906">
  <param name="_ExtentY" value="8996">
  <param name="AUTOSTART" value="-1">
  <param name="SHUFFLE" value="0">
  <param name="PREFETCH" value="0">
  <param name="NOLABELS" value="0">
  <param name="SRC" value="<%response.write url_true(upload_file,ispic)%>">
  <param name="CONTROLS" value="ImageWindow">
  <param name="CONSOLE" value="Clip1">
  <param name="LOOP" value="0">
  <param name="NUMLOOP" value="0">
  <param name="CENTER" value="0">
  <param name="MAINTAINASPECT" value="0">
  <param name="BACKGROUNDCOLOR" value="#000000">
  <embed  type="audio/x-pn-realaudio-plugin" console="Clip1" controls="ImageWindow" autostart="false"></embed> 
</object>
<object id="video1" classid="clsid:CFCDAA03-8BE4-11cf-B84B-0020AFBBCCFA" width=450 height=60>
  <param name="_ExtentX" value="11906">
  <param name="_ExtentY" value="1588">
  <param name="AUTOSTART" value="-1">
  <param name="SHUFFLE" value="0">
  <param name="PREFETCH" value="0">
  <param name="NOLABELS" value="0">
  <param name="CONTROLS" value="ControlPanel,StatusBar">
  <param name="CONSOLE" value="Clip1">
  <param name="LOOP" value="0">
  <param name="NUMLOOP" value="0">
  <param name="CENTER" value="0">
  <param name="MAINTAINASPECT" value="0">
  <param name="BACKGROUNDCOLOR" value="#000000">
  <embed type="audio/x-pn-realaudio-plugin" console="Clip1" controls="ControlPanel,StatusBar" width=450 height=60 autostart="false"></embed>
</object>

<% case "asf","wmv","mpg","mpeg","wma","asx","mms","avi" %>
<object id="beyondest.com.mPlayer" width=544 height=440 classid="CLSID:6BF52A52-394A-11D3-B153-00C04F79FAA6" type="application/x-oleobject" standby="正在载入 Windows Media Player 播放流 ...">
  <param name="URL" value="file/video/<%response.write url_true(upload_file,ispic)%>">
  <param name="Album" value="Beyondest.com"/>
  <param name="rate" value="1">
  <param name="balance" value="0">
  <param name="currentPosition" value="0">
  <param name="defaultFrame" value="">
  <param name="playCount" value="100">
  <param name="autoStart" value="-1">
  <param name="currentMarker" value="0">
  <param name="invokeURLs" value="-1">
  <param name="baseURL" value="">
  <param name="volume" value="100">
  <param name="mute" value="0">
  <param name="uiMode" value="full">
  <param name="stretchToFit" value="0">
  <param name="windowlessVideo" value="0">
  <param name="enabled" value="-1">
  <param name="enableContextMenu" value="0">
  <param name="fullScreen" value="0">
  <param name="SAMIStyle" value="">
  <param name="SAMILang" value="">
  <param name="SAMIFilename" value="">
  <param name="captioningID" value="">
</object>
<%
  end select
end sub

sub gallery_main(gma)
  dim j,k,kn,pic,name,nnum,pic_link,pic_temp,ntypes:nnum=1
  pageurl="?action="&action&"&"
  if cid>0 then
    sqladd=" and c_id="&cid
    pageurl=pageurl&"c_id="&cid&"&"
    if sid>0 then
      sqladd=sqladd&" and s_id="&sid
      pageurl=pageurl&"s_id="&sid&"&"
    end if
  end if

  sql="select id,c_id,s_id,types,spic,pic,name,counter,power,emoney,remark,username from gallery where hidden=1 and types='"&action&"'"&sqladd&" order by id desc"
  set rs=server.createobject("adodb.recordset")
  rs.open sql,conn,1,1
  if not(rs.eof and rs.bof) then rssum=rs.recordcount
  
  select case action
  case "logo"
    nummer=30
  case "baner"
    nummer=6
  case "film"
    nummer=6
  case else
    nummer=9
  end select
  call format_pagecute()
%>
<table border=0 cellspacing=0 cellpadding=0 width='100%' height=35>
<tr>
<td background='images/main/bar_3_bg.gif' width=35 valign=bottom><img border='0' src='images/main/icon_1.gif'></td>
<td background='images/main/bar_3_bg.gif' valign=top>
<table border=0 cellspacing=0 cellpadding=0 width='100%' height=30><tr><td valign=middle>&nbsp;<a href='news.asp'><b><font class=end><%if tit2="相册" then tit2="照片"
response.Write tit2%>列表</font></b></a></td></tr></table>
</td>
<td valign=top  width=30 background='images/main/bar_3_bg.gif' align=right><img border=0 src='images/main/bar_1_rt.gif'></td>
</tr></table>

<table border=0 width='100%'>
<%
select case action
case "logo"
  kn=5:nummer=nummer*2
  if nummer mod kn > 0 then
    k=nummer\kn+1
  else
    k=nummer\kn
  end if
  
  if int(viewpage)>1 then
    rs.move (viewpage-1)*nummer
  end if
  
  for i=1 to k
    'if rs.eof then exit for
    response.write "<tr align=center>"
    for j=1 to kn
      if rs.eof or nnum>nummer then exit for
      pic=rs("pic"):name=rs("name")
      response.write "<td><table border=0><tr><td align=center><img src='images/"&pic&"' border=0 width=88 height=31></td></tr><tr><td align=center><font title='"&name&"'>"&code_html(name,1,10)&"</font></td></tr></table></td>"
      rs.movenext
      nnum=nnum+1
    next
    response.write "</tr>"
  next

case "baner"
  nummer=6
  if int(viewpage)>1 then
    rs.move (viewpage-1)*nummer
  end if
  for i=1 to nummer
    if rs.eof then exit for
    pic=rs("pic"):name=rs("name")
    response.write "<tr><td align=center><a href='images/photo/"&pic&".jpg' target='_blank'><img src='images/photo/"&pic&".jpg' border=0 onload='javascript:if(this.width>500)this.width=500'></a></td></tr><tr><td align=center>"&kong&"</td></tr>"
    rs.movenext
  next

case "film"
  pic_width=200
  pic_height=150
  kn=3
  if nummer mod kn > 0 then
    k=nummer\kn+1
  else
    k=nummer\kn
  end if
  if int(viewpage)>1 then
    rs.move (viewpage-1)*nummer
  end if
  for i=1 to nummer
      if rs.eof or nnum>nummer then exit for
      if action="paste" then
        pic=rs("pic")
      else
        pic=rs("spic")
      end if
      name=rs("name")
      if action<>"paste" then
        pic_link="?types=view&action="&action&"&c_id="&rs("c_id")&"&s_id="&rs("s_id")&"&id="&rs("id")
      else
        pic_link=web_var(web_down,5)&"/"&pic
      end if
      pic_temp="<table border=0>" & _
	       "<tr><td align=center width="&pic_width+10&"  valign=top><a href='"&pic_link&"' target='_blank'><img src='images/video/"&pic&"' border=0 width="&pic_width&" height="&pic_height&"></a></td>"
      if action<>"paste" then
	pic_temp=pic_temp&"<td align=left width="&-pic_width+540&">"&kong&"<b><font class=big>"&name&"</font></b>"&kong&"点击：<font class=red>"&rs("counter")&"</font>次&nbsp;&nbsp;┋&nbsp;&nbsp;权限：注册用户&nbsp;&nbsp;┋&nbsp;&nbsp;整理："&format_user_view(rs("username"),1,"")&kong&"说明："&rs("remark")&"</td>" 
      end if
      pic_temp=pic_temp&"</tr></table>"
      response.write "<tr ><td align=center>"&format_k(pic_temp,1,5,550,pic_height+10)&"</td></tr>"
     rs.movenext
  next


case else
  kn=3
  if nummer mod kn > 0 then
    k=nummer\kn+1
  else
    k=nummer\kn
  end if
  if int(viewpage)>1 then
    rs.move (viewpage-1)*nummer
  end if
  for i=1 to k
    'if rs.eof then exit for
    response.write "<tr><td height=5></td></tr><tr align=center>"
    for j=1 to kn
      if rs.eof or nnum>nummer then exit for
      if action="paste" then
        pic=rs("pic")
      else
        pic=rs("spic")
      end if
      name=rs("name")
      if action<>"paste" then
        pic_link="?types=view&action="&action&"&c_id="&rs("c_id")&"&s_id="&rs("s_id")&"&id="&rs("id")
      else
        pic_link=web_var(web_down,5)&"/"&pic
      end if
      pic_temp="<table border=0>" & _
	       "<tr><td align=center><a href='"&pic_link&"' target='_blank'><img src='images/"&pic&"' border=0 width="&pic_width&" height="&pic_height&"></a></td></tr>"
	       
      if action<>"paste" then
	pic_temp=pic_temp&"<tr><td align=center class=blue><b>"&code_html(name,1,20)&"</b></td></tr><tr><td align=center>权限:<font class=red_3>注册用户</font>&nbsp;&nbsp;点击:<font class=red>"&rs("counter")&"次</font></td></tr>" 
      end if
      pic_temp=pic_temp&"</table>"
      response.write "<td height="&pic_height+50&">"&format_k(pic_temp,1,5,pic_width+10,pic_height+30)&"</td>"
      rs.movenext
      nnum=nnum+1
    next
    response.write "</tr>"
  next
end select

  rs.close:set rs=nothing
%>
<tr><td height=1 colspan=3 background='images/bg_dian.gif'></td></tr>
<tr><td align=center height=30 colspan=3>

<table border=0 width='100%' cellspacing=0 cellpadding=0>
<tr align=center valign=bottom><td width='40%' align=left>
现有<font class=red><%response.write rssum%></font>个文件┋
每页<font class=red><%response.write nummer%></font>个
</td><td width='60%' align=right>
页次：<font class=red><%response.write viewpage%></font>/<font class=red><%response.write thepages%></font> 分页：<%response.write jk_pagecute(nummer,thepages,viewpage,pageurl,5,"#ff0000")%>
</td></tr>
</table>


</td></tr>
</table>
<%
end sub
sub gallery_view()
  dim ispic,width,height
  width=web_var(web_num,9):height=web_var(web_num,10)
  sql="select * from gallery where hidden=1 and types='"&action&"' and id="&id
  set rs=conn.execute(sql)
  if rs.eof and rs.bof then rs.close:call gallery_main(action):exit sub
  ispic=rs("pic")
  if len(ispic)<1 then rs.close:call gallery_main(action):exit sub
  
  if action<>"paste" then call emoney_notes(rs("power"),rs("emoney"),n_sort,id,"js",0,1,"gallery.asp?action="&action&"&c_id="&cid&"&s_id="&sid)
  sql="update gallery set counter=counter+1 where id="&id
  conn.execute(sql)
%>
<table border=0 width='98%' align=center class=tf>
<tr><td height=30 align=center><font class=blue size='4'><b><%response.write code_html(rs("name"),1,0)%></b></font></td></tr>
<tr><td height=5></td></tr>
<tr><td align=center class=bw>
<%
  select case action
  case "flash"
  upload_file=web_var(web_down,5)
%>
<object classid="clsid:D27CDB6E-AE6D-11cf-96B8-444553540000" codebase="http://download.macromedia.com/pub/shockwave/cabs/flash/swflash.cab#version=5,0,0,0" width="<%response.write width%>" height="<%response.write height%>">
<param name=movie value=<%response.write url_true(upload_file,ispic)%>>
<param name=quality value=high>
<embed src="<%response.write url_true(upload_file,ispic)%>" quality=high pluginspage="http://www.macromedia.com/shockwave/download/index.cgi?P1_Prod_Version=ShockwaveFlash" type="application/x-shockwave-flash" width="<%response.write width%>" height="<%response.write height%>"></embed> 
</object>
<%
  case "film"
    call film_view(ispic,width,height)
  case else
%>
<img src='<%response.write url_true(upload_file,ispic)%>' border=0>
<% end select %>
</td></tr>
<tr><td height=5></td></tr>
</table>
<table border=0 width='98%' align=center class=tf>
<tr><td>文件说明：<%response.write code_jk(rs("remark"))%></td></tr>
<tr><td>上传用户：<%response.write format_user_view(rs("username"),1,"")%>　　上传时间：<%response.write rs("tim")%>　　人气：<font class=red><%response.write rs("counter")%></font>　　<%response.write img_small("jt0")%><a href='<%response.write url_true(upload_file,ispic)%>' target=_blank>在新窗口中浏览</a></td></tr>
<tr><td height=10></td></tr>
<tr><td align=center><% call review_type(n_sort,id,"gallery.asp?action=view&c_id="&cid&"&s_id="&sid&"&id="&id,1) %></td></tr>
</table>
<%
end sub

%>
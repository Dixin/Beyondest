<%
' ====================
' Beyondest.Com v3.6.1
' http://beyondest.com
' ====================

dim n_sort,cid,sid
cid=0:sid=0

function class_sort(sns,surl,idc,ids)
  dim nid,rs,sql,rs2,sql2,nnid,j,k:k=12
  class_sort=vbcrlf&"<table border=0><tr><td height=3></td></tr>"
  sql="select c_name,c_id from jk_class where nsort='"&sns&"' order by c_order"
  set rs=conn.execute(sql)
  do while not rs.eof
    nid=rs("c_id")
    class_sort=class_sort&vbcrlf&"<tr valign=top><td><table border=0><tr><td height=18>|&nbsp;<a href='"&surl&"_list.asp?c_id="&nid&"'"
    if idc=nid then class_sort=class_sort&" class=red"
    class_sort=class_sort&">"&rs("c_name")&"</a>&nbsp;|</td></tr></table></td><td>"&vbcrlf&"<table border=0>"
    sql2="select s_name,s_id from jk_sort where c_id="&nid&" order by s_order"
    set rs2=conn.execute(sql2)
    do while not rs2.eof
      class_sort=class_sort&"<tr>"
      for j=1 to k
        if rs2.eof then exit for
        nnid=rs2("s_id")
        class_sort=class_sort&vbcrlf&"<td>&nbsp;<a href='"&surl&"_list.asp?c_id="&nid&"&s_id="&rs2("s_id")&"'"
        if ids=nnid then class_sort=class_sort&" class=red_3"
        class_sort=class_sort&">"&rs2("s_name")&"</a>&nbsp;</td>"
        rs2.movenext
      next
      class_sort=class_sort&"</tr>"
    loop
    class_sort=class_sort&"</table></td></tr>"
    rs.movenext
  loop
  rs.close:set rs=nothing
  class_sort=class_sort&vbcrlf&"<tr><td height=1></td></tr></table>"
end function


function class_sortp(sns,surl,idc,ids)
  dim nid,rs,sql,rs2,sql2,stt,con,spic,nnid,j,k:k=7
  class_sortp=""
  sql="select c_name,c_id from jk_class where nsort='"&sns&"' order by c_order"
  set rs=conn.execute(sql)
  do while not rs.eof
    nid=rs("c_id")
    stt="&nbsp;<a href='"&surl&"_list.asp?c_id="&nid&"'"
    if idc=nid then stt=stt&" class=red"
    stt=stt&">"&rs("c_name")&"</a>"
    sql2="select s_name,s_id,pic from jk_sort where c_id="&nid&" order by s_order"
    set rs2=conn.execute(sql2)
    con="<table wdth=100% border=0 cellspacing=0 cellpadding=0><tr align=left>"
    do while not rs2.eof
        nnid=rs2("s_id")
        spic=rs2("pic")
        con=con&"<td width=113 valign=left>"&kong&"<table border=0 cellspacing=0 cellpadding=0><tr><td align=left>&nbsp;<a href='"&surl&"_list.asp?c_id="&nid&"&s_id="&rs2("s_id")&"'"
        if ids=nnid then con=con&" class=red_3"
        con=con&"><img src=images/down/"&spic&"x.jpg width=80 height=80 border=0></a></td><tr><td height=20 align=left>&nbsp;<a href='"&surl&"_list.asp?c_id="&nid&"&s_id="&rs2("s_id")&"'"
        if ids=nnid then con=con&" class=red_3"
        con=con&">"&rs2("s_name")&"</a>&nbsp;</td></tr></table></td><td width=1 bgcolor="&web_var(web_color,3)&"></td>"
        rs2.movenext
    loop
    con=con&"<td></td></tr></table>"
    class_sortp=class_sortp&format_barc(stt,con,5,0,71)
    rs.movenext
  loop
  rs.close:set rs=nothing
end function

function nsort_left(n_sort,cc_id,ss_id,link_url,left_type)
  dim rs1,sql1,rs2,sql2,ccid,ssid:cc_id=int(cc_id):ss_id=int(ss_id)
  nsort_left=vbcrlf&"<table border=0><tr><td height=1></td></tr>"
  sql1="select c_id,c_name from jk_class where nsort='"&n_sort&"' order by c_order,c_id"
  set rs1=conn.execute(sql1)
  do while not rs1.eof
    ccid=int(rs1(0))
    if cc_id=ccid or left_type=0 then
      nsort_left=nsort_left&vbcrlf&"<tr><td>"
      if cc_id=ccid then
        nsort_left=nsort_left&img_small("jt1")
      else
        nsort_left=nsort_left&img_small("jt12")
      end if
      nsort_left=nsort_left&"<a href='"&link_url&"c_id="&ccid&"'>"&rs1(1)&"</a></td></tr>"
      sql2="select s_id,s_name from jk_sort where c_id="&ccid&" order by s_order,s_id"
      set rs2=conn.execute(sql2)
      do while not rs2.eof
        ssid=int(rs2(0))
        nsort_left=nsort_left&vbcrlf&"<tr><td>&nbsp;&nbsp;"&img_small("jt0")&"<a href='"&link_url&"c_id="&ccid&"&s_id="&rs2(0)&"'"
        if ssid=ss_id then nsort_left=nsort_left&" class=red_3"
        nsort_left=nsort_left&">"&rs2(1)&"</a></td></tr>"
        rs2.movenext
      loop
      rs2.close:set rs2=nothing
    else
      nsort_left=nsort_left&vbcrlf&"<tr><td>"&img_small("jt12")&"<a href='"&link_url&"c_id="&ccid&"'>"&rs1(1)&"</a></td></tr>"
    end if
    rs1.movenext
  loop
  rs1.close:set rs1=nothing
  nsort_left=nsort_left&vbcrlf&"<tr><td height=1></td></tr></table>"
end function

sub cid_sid_sql(csst,csstt)
  dim temp:temp=csstt
  if cid>0 then
    sqladd=sqladd&" and c_id="&cid
    pageurl=pageurl&"c_id="&cid&"&"
    if sid>0 then
      sqladd=sqladd&" and s_id="&sid
      pageurl=pageurl&"s_id="&sid&"&"
    end if
  end if
  if csst=1 or csst=2 then
    if len(keyword)>0 then
      sqladd=sqladd&" and "&temp&" like '%"&keyword&"%'"
      pageurl=pageurl&"keyword="&server.urlencode(keyword)&"&"
      if csst=2 then pageurl=pageurl&"sea_type="&sea_type&"&"
    end if
  end if
end sub

sub cid_sid()
  cid=trim(request.querystring("c_id"))
  sid=trim(request.querystring("s_id"))
  if not(isnumeric(cid)) then
    if len(cid)>0 then
      cid=replace(cid,"&s_id=",",")
      sid=mid(cid,instr(cid,",")+1,len(cid))
      cid=mid(cid,1,instr(cid,",")-1)
    else
      cid=0
    end if
  end if
  if not(isnumeric(cid)) then cid=0
  if not(isnumeric(sid)) then sid=0
  cid=int(cid):sid=int(sid)
end sub

function put_type(pts)
  put_type=""
  select case pts
  case "article"
    put_type="我要发表文章"
  case "news"
    put_type="我要发布新闻"
  case "down"
    put_type="我要添加音乐"
  case "website"
    put_type="我要推荐网站"
  case "gallery"
    put_type="我要上传贴图"
  end select
  if put_type<>"" then put_type="[ <a href='user_put.asp?action="&pts&"'>→ "&put_type&"</a> ]"
end function
%>
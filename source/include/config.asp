<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
<%
Option Explicit
'Response.buffer=true 		'开启缓冲页面功能，只要把前面的引号去掉就可以开启了

dim web_config,web_cookies,web_login,web_setup,web_menu,web_num,web_color,web_upload,web_safety,web_edition,web_label,web_lct,web_imglct
dim web_news_art,web_down,web_shop,web_stamp,space_mod,gang,gang2,sk_bar_lf
dim web_error,user_grade,user_power,forum_type,now_time,timer_start,redx,kong,ukong,go_back,closer,lefter,righter
dim login_username,login_password,login_mode,login_message,login_popedom,login_emoney
dim rs,sql,i,tit,tit_fir,index_url,action,page_power,m_unit,sk_bar,sk_class,sk_jt,sk_img
%>
<!-- #include file="common.asp" -->
<!-- #include file="functions.asp" -->
<%
sk_bar=11:sk_bar_lf=15:sk_class="end":sk_jt="jt0":space_mod=web_var(web_num,12):m_unit=web_var(web_config,8):now_time=time_type(now(),9):timer_start=timer()
web_label="<a href='http://www.beyondest.com/' target=_blank>Power by <b><font face=Arial color=#CC3300>Beyondest</font><font face=Arial>.Com</font></b></a>"
redx="&nbsp;<font color='#ff0000'>*</font>&nbsp;":kong="<table width='100%' height=2><tr><td></td></tr></table>":gang="<table height=1 bgcolor="&web_var(web_color,3)&" width='100%' cellspacing=0 cellpadding=0 border=0><tr><td></td></tr></table>":gang2="<table width=1 height='100%' bgcolor="&web_var(web_color,3)&"><tr><td></td></tr></table>":web_edition="Beyondest V3.6 Demo":ukong="<table border=0><tr><td height=6></td></tr></table>"
go_back="<a href='javascript:history.back(1)'>返回上一页</a>":closer="<a href='javascript:self.close()'>『关闭窗口』</a>"
login_mode="":login_popedom="":login_message=0
login_username=trim(request.cookies(web_cookies)("login_username"))
login_password=trim(request.cookies(web_cookies)("login_password"))
action=trim(request.querystring("action"))
sk_img="&nbsp;<img src='images/small/img.gif' border=0>"

function format_barc(b_tt,b_ct,b_tp,b_cl,b_ic)
    dim tempbc,bheight1,bheight2,bheight3,bheight5
    bheight1=35:bheight2=30:bheight3=25:bheight5=27
    tempbc="<table border=0 cellspacing=0 cellpadding=0 width='100%'><tr><td>"
    select case b_tp
        case 1
          tempbc=tempbc&"<table border=0 cellspacing=0 cellpadding=0 width='100%' height="&bheight1&"><tr><td background='images/"&web_var(web_config,5)&"/bar_1_bg.gif' width="&bheight1&"valign=bottom>"&format_icon(b_ic)&"</td><td background='images/"&web_var(web_config,5)&"/bar_1_bg.gif' valign=top><table border=0 cellspacing=0 cellpadding=0 width='100%' height="&bheight2&"><tr><td valign=middle>&nbsp;"&b_tt&"</td></tr></table></td><td valign=top  width="&bheight2&" background='images/"&web_var(web_config,5)&"/bar_1_bg.gif' align=right>"&format_img("bar_1_rt.gif")&"</td></tr></table>"
        case 2
          tempbc=tempbc&"<table border=0 cellspacing=0 cellpadding=0 width='100%' height="&bheight1&"><tr><td background='images/"&web_var(web_config,5)&"/bar_2_bg.gif' width="&bheight1&"valign=top>"&format_icon(b_ic)&"</td><td background='images/"&web_var(web_config,5)&"/bar_2_bg.gif' valign=bottom><table border=0 cellspacing=0 cellpadding=0 width='100%' height="&bheight2&"><tr><td valign=middle>&nbsp;"&b_tt&"</td></tr></table></td><td valign=bottom  width="&bheight2&" background='images/"&web_var(web_config,5)&"/bar_2_bg.gif' align=right>"&format_img("bar_2_rt.gif")&"</td></tr></table>"
        case 3
          tempbc=tempbc&"<table border=0 cellspacing=0 cellpadding=0 width='100%' height="&bheight1&"><tr><td background='images/"&web_var(web_config,5)&"/bar_3_bg.gif' width="&bheight1&"valign=bottom>"&format_icon(b_ic)&"</td><td background='images/"&web_var(web_config,5)&"/bar_3_bg.gif' valign=top><table border=0 cellspacing=0 cellpadding=0 width='100%' height="&bheight2&"><tr><td valign=middle>&nbsp;"&b_tt&"</td></tr></table></td><td valign=top  width="&bheight2&" background='images/"&web_var(web_config,5)&"/bar_3_bg.gif' align=right>"&format_img("bar_1_rt.gif")&"</td></tr></table>"
        case 4
          tempbc=tempbc&"<table border=0 cellspacing=0 cellpadding=0 width='100%' height="&bheight5&" background='images/"&web_var(web_config,5)&"/bar_4_bg.gif'><tr><td width="&bheight1&"valign=bottom>"&format_img("bar_4_lf.gif")&"</td><td><table border=0 cellspacing=0 cellpadding=0 width='100%' height="&bheight5&"><tr><td valign=middle>&nbsp;"&b_tt&"</td></tr></table></td></tr></table>"
        case 5
          tempbc=tempbc&"<table border=0 cellspacing=0 cellpadding=0 width='100%' height="&bheight5&"><tr><td background='images/"&web_var(web_config,5)&"/bar_3_bg.gif' width="&bheight1&"valign=bottom>"&format_icon(b_ic)&"</td><td background='images/"&web_var(web_config,5)&"/bar_3_bg.gif' valign=top><table border=0 cellspacing=0 cellpadding=0 width='100%' height="&bheight5&"><tr><td valign=middle>&nbsp;"&b_tt&"</td></tr></table></td><td valign=top  width="&bheight2&" background='images/"&web_var(web_config,5)&"/bar_3_bg.gif' align=right>"&format_img("bar_5_rt.gif")&"</td></tr></table>"
    end select
    select case b_cl
        case 0
           tempbc=tempbc&"</td></tr><tr><td>"&b_ct&"</td></tr></table>"
        case 1
          tempbc=tempbc&"</td></tr><tr><td bgcolor="&web_var(web_color,1)&">"&b_ct&ukong&"</td></tr></table>"
    end select
    format_barc=tempbc
end function

function format_icon(icon_icon)
    format_icon="<img border='0' src='images/"&web_var(web_config,5)&"/icon_"&icon_icon&".gif'>"
end function

function format_bar(bar_var,bar_body,bar_type,bar_fk,bar_jt,bar_color,bar_more)
  dim bar_temp,bar_vars,bar_mores,bar_height
  bar_height=30:bar_mores=""
  bar_vars="<table border=0 cellspacing=0 cellpadding=0><tr><td>&nbsp;"
  if isnumeric(bar_jt) then
    if bar_jt<>0 then bar_vars=bar_vars&"<img border=0 src='images/"&web_var(web_config,5)&"/bar_"&bar_type&"_jt.gif' align=absmiddle>&nbsp;"
  else
    if bar_jt<>"" then bar_vars=bar_vars&img_small(bar_jt)
  end if
  bar_vars=bar_vars&bar_var&"</td></tr></table>"
  if bar_more<>"" then bar_mores=bar_more&"&nbsp;&nbsp;"
  
  bar_temp=vbcrlf&"<table border=0 width='100%' cellspacing=0 cellpadding=0"
  select case int(left(bar_type,1))
  case 0
    bar_temp=bar_temp&"><tr>" & _
	     vbcrlf&"<td>"&bar_vars&"</td>" & _
	     vbcrlf&"<td align=right>"&bar_mores&"</td>"
  case 1
    bar_temp=bar_temp&"><tr>" & _
	     vbcrlf&"<td width=30 valign=top>"&format_img("bar_"&bar_type&"_left.gif")&"</td>" & _
	     vbcrlf&"<td background='images/"&web_var(web_config,5)&"/bar_"&bar_type&"_bg.gif'><table border=0 width='100%' cellspacing=0 cellpadding=0><tr><td>"&bar_vars&"</td><td align=right>"&bar_more&"</td></tr></table></td>" & _
	     vbcrlf&"<td width=20>"&format_img("bar_"&bar_type&"_right.gif")&"</td>"
  case 2
    bar_temp=bar_temp&"><tr>" & _
	     vbcrlf&"<td width=30 valign=top>"&format_img("bar_"&bar_type&"_left.gif")&"</td>" & _
	     vbcrlf&"<td width="&web_var(bar_color,3)&" background='images/"&web_var(web_config,5)&"/bar_"&bar_type&"_bg0.gif'>"&bar_vars&"</td>" & _
	     vbcrlf&"<td width=20>"&format_img("bar_"&bar_type&"_center.gif")&"</td>" & _
	     vbcrlf&"<td background='images/"&web_var(web_config,5)&"/bar_"&bar_type&"_bg.gif' align=right>&nbsp;"&bar_more&"</td>" & _
	     vbcrlf&"<td width=20 align=right>"&format_img("bar_"&bar_type&"_right.gif")&"</td>"
  end select
  bar_temp=bar_temp&vbcrlf&"</tr></table>"
  
  if bar_fk=1 or bar_fk=3 then
    bar_body="<table border=0 width='98%' cellspacing=4 cellpadding=4><tr><td>"&bar_body&"</td></tr></table>"
  else
    bar_body="<table border=0 width='100%' cellspacing=0 cellpadding=0><tr><td>"&bar_body&"</td></tr></table>"
  end if
  format_bar="<table width='100%' cellspacing=0 cellpadding=0"
  select case bar_fk
  case 0,1
    format_bar=format_bar&" border=0>" & _
	       vbcrlf&"<tr><td height="&bar_height&" valign=bottom"
    if int(left(bar_type,1))=0 then format_bar=format_bar&"bgcolor="&web_var(bar_color,1)
    format_bar=format_bar&" background='"&web_var(bar_color,3)&"'>"&bar_temp&"</td></tr>" & _
	       vbcrlf&"<tr><td align=center"
    if web_var(bar_color,2)<>"" then format_bar=format_bar&" bgcolor="&web_var(bar_color,2)
    format_bar=format_bar&">"&bar_body&"</td></tr></table>"
  case 2,3
    if int(left(bar_type,1))=0 then
      format_bar=format_bar&" border=1 bgcolor="&web_var(bar_color,2)&" bordercolor="&web_var(bar_color,1)&">" & _
		 vbcrlf&"<tr><td height="&bar_height&" bgcolor="&web_var(bar_color,1)&" background='"&web_var(bar_color,3)&"' valign=bottom>"&bar_temp&"</td></tr>" & _
		 vbcrlf&"<tr><td align=center bordercolor="&web_var(bar_color,2)&">"&bar_body&"</td></tr>"&vbcrlf&"</table>"
    else
      format_bar=format_bar&" border=0>" & _
		 vbcrlf&"<tr><td height="&bar_height&" valign=bottom>"&bar_temp&"</td></tr><tr><td align=center>" & _
		 vbcrlf&"<table border=0 width='100%' cellspacing=0 cellpadding=0><tr align=center><td width=1 bgcolor="&web_var(bar_color,1)&"></td><td bgcolor="&web_var(bar_color,2)&">"&bar_body&"</td><td width=1 bgcolor="&web_var(bar_color,1)&"></td></tr><tr><td height=1 colspan=3 bgcolor="&web_var(bar_color,1)&"></td></tr></table>" & _
		 vbcrlf&"</td></tr></table>"
    end if
  end select
end function


function format_table(btype,tc)
  'response.write tc
  select case btype
  case 1
    format_table="<table border=0 width='98%' cellspacing=1 cellpadding=4 bgcolor="&web_var(web_color,tc)&" bordercolor="&web_var(web_color,1)&">"
  case 2
    format_table=""
  case 3
    format_table=" valign=middle bgcolor="&web_var(web_color,tc)&" bordercolor="&web_var(web_color,tc)
  case 4
    format_table=" background='images/"&web_var(web_config,5)&"/bg_table.gif' bordercolor="&web_var(web_color,tc)
  end select
end function

function format_k(kvar,kt,kk,kw,kh)
  dim temp1,t1
  t1="images/"&web_var(web_config,5)&"/k"&kt&"_"
  temp1=vbcrlf&"<table border=0 width="&kw+kk*2&" height="&kh+kk*2&" cellpadding=0 cellspacing=0>" & _
	vbcrlf&"<tr>" & _
	vbcrlf&"<td width="&kk&" height="&kk&"><img src='"&t1&"1.gif' border=0></td>" & _
	vbcrlf&"<td width="&kw&" height="&kk&" background='"&t1&"top.gif'></td>" & _
	vbcrlf&"<td width="&kk&" height="&kk&"><img src='"&t1&"2.gif' border=0></td>" & _
	vbcrlf&"</tr>" & _
	vbcrlf&"<tr>" & _
	vbcrlf&"<td width="&kk&" height="&kh&" background='"&t1&"left.gif'></td>" & _
	vbcrlf&"<td width="&kw&" height="&kh&" align=center>"&kvar&"</td>" & _
	vbcrlf&"<td width="&kk&" height="&kh&" background='"&t1&"right.gif'></td>" & _
	vbcrlf&"</tr>" & _
	vbcrlf&"<tr>" & _
	vbcrlf&"<td width="&kk&" height="&kk&"><img src='"&t1&"3.gif' border=0></td>" & _
	vbcrlf&"<td width="&kw&" height="&kk&" background='"&t1&"end.gif'></td>" & _
	vbcrlf&"<td width="&kk&" height="&kk&"><img src='"&t1&"4.gif' border=0></td>" & _
	vbcrlf&"</tr>" & _
	vbcrlf&"</table>"
  format_k=temp1
end function
sub format_pagecute()
  if rssum mod nummer > 0 then
    thepages=rssum\nummer+1
  else
    thepages=rssum\nummer
  end if
  page=trim(request("page"))
  if not(isnumeric(page)) then page=1
  if int(page)>int(thepages) or int(page)<1 then
    viewpage=1
  else
    viewpage=int(page)
  end if
end sub
function format_menu(mvars)
  dim i,mdim,mvar:mvar=trim(mvars):format_menu=""
  mdim=split(web_menu,"|")
  for i=0 to ubound(mdim)
    if mvar=left(mdim(i),instr(mdim(i),":")-1) then format_menu=right(mdim(i),len(mdim(i))-instr(mdim(i),":")):exit for
  next
  erase mdim
end function
function format_user_power(uname,umode,pvar)
  dim admint:admint=format_power2(1,1):format_user_power="yes"
  if umode=admint then exit function
  if instr("|"&pvar&"|","|"&uname&"|")<1 then format_user_power="no"
end function
function format_page_power(umode)
  dim unum:unum=format_power(umode,2):format_page_power="yes"
  if page_power="" then exit function
  if instr("."&page_power&".","."&unum&".")<1 then format_page_power="no"
end function
sub user_integral(ut,unum,uuser)
  dim fh:fh="+"
  if ut="del" then fh="-"
  conn.execute("update user_data set integral=integral"&fh&unum&" where username='"&uuser&"'")
end sub
function format_power(pvar,pt)
  dim i,pdim:pvar=trim(pvar)
  if pt=2 then
    format_power=0
  else
    format_power=""
  end if
  pdim=split(user_power,"|")
  for i=0 to ubound(pdim)
    if pvar=left(pdim(i),instr(pdim(i),":")-1) then
      select case pt
      case 1
        format_power=right(pdim(i),len(pdim(i))-instr(pdim(i),":")):exit for
      case 2
        format_power=i+1:exit for
      case else
        format_power=pvar:exit for
      end select
    end if
  next
  erase pdim
end function
function format_power2(pnn,pt)
  dim i,pdim,pn:format_power2="":pn=pnn-1
  pdim=split(user_power,"|")
  if pn<=ubound(pdim) then
    if pt=1 then
      format_power2=left(pdim(pn),instr(pdim(pn),":")-1)
    else
      format_power2=right(pdim(pn),len(pdim(pn))-instr(pdim(pn),":"))
    end if
  end if
  erase pdim
end function
function power_pic(emon,pp,pt)
  power_pic=""
  if pt=1 then power_pic="<font class=red_3>免费下载</font>&nbsp;&nbsp;&nbsp;"
  dim ddim,j:ddim=split(pp,".")
  for j=0 to ubound(ddim)
    if int(ddim(j))=0 then
      power_pic=power_pic&img_small("icon_other")
    else
      power_pic=power_pic&img_small("icon_"&format_power2(ddim(j),1))
    end if
  next
  erase ddim
end function
function user_star(u_s,u_p,u_t)
  dim tempp,tempn,ui,sdim,sn,u1,u2,uu
  tempp="":tempn="":u_s=int(u_s):u_t=int(u_t)
  select case u_p
  case format_power2(1,1)
    user_star=format_power2(1,u_t):exit function
  case format_power2(2,1)
    user_star=format_power2(2,u_t):exit function
  case format_power2(3,1)
    tempp="p"
  end select
  sdim=split(user_grade,"|"):sn=ubound(sdim)
  for ui=0 to sn
    u1=int(left(sdim(ui),instr(sdim(ui),":")-1))
    select case ui
    case 0
      u1=int(left(sdim(ui+1),instr(sdim(ui+1),":")-1))
      if u_s<u1 then
        uu=ui:exit for
      elseif u_s=u1 then
        uu=ui+1:exit for
      end if
    case sn
      response.write u_s&"-"&u1
      if u_s>=u1 then uu=ui:exit for
    case else
      u2=int(left(sdim(ui+1),instr(sdim(ui+1),":")-1))
      if u_s>=u1 and u_s<u2 then uu=ui:exit for
    end select
  next
  if u_t=2 then
    tempp=right(sdim(uu),len(sdim(uu))-instr(sdim(uu),":"))
  else
    tempp=tempp&uu
  end if
  erase sdim:user_star=tempp
end function
function user_power_type(ptt)
  dim pdim,pn:user_power_type="网站用户图例："
  pdim=split(user_power,"|")
  for pn=0 to ubound(pdim)
    user_power_type=user_power_type&"&nbsp;"&img_small("icon_"&left(pdim(pn),instr(pdim(pn),":")-1))&right(pdim(pn),len(pdim(pn))-instr(pdim(pn),":"))
  next
  erase pdim
  user_power_type=user_power_type&"&nbsp;&nbsp;"&img_small("icon_other")&"游客"
end function
function popedom_format(popedom_var,popedom_n)
  dim poptemp:poptemp=0
  if len(popedom_var)=50 then poptemp=int(mid(popedom_var,popedom_n,1))
  popedom_format=poptemp
end function
sub emoney_notes(power,emoney,n_sort,iid,err_type,rss,conns,url)
  dim temp1:temp1=emoney_note(power,emoney,n_sort,iid)
  if temp1<>"yes" then
    if int(rss)=1 then rs.close:set rs=nothing
    if int(conns)=1 then call close_conn()
    select case err_type
    case "error"
      call cookies_type("power")
    case "js"
%><script language=javascript>
alert("您没有足够的权限进行刚才的操作！\n\n点击返回……");
location.href='<%response.write url%>';
</script><%
      response.end
    end select
  end if
end sub
function emoney_note(power,emoney,n_sort,iid)
  dim userp,sql,rs,notess:notess="no"
  if len(power)>0 then
    userp=format_power(login_mode,2)
    if not(isnumeric(userp)) then userp=0
    userp=int(userp)
    if userp=0 then login_emoney=0
    if instr(1,"."&power&".","."&userp&".")>0 then notess="yes"
  end if
  if login_mode="" and int(emoney)>0 then notess="no"
  if notess="yes" then
    set rs=conn.execute("select id from notes where username='"&login_username&"' and nsort='"&n_sort&"' and iid="&iid)
    if rs.eof and rs.bof then notess="no2"
    rs.close:set rs=nothing
  end if
  if notess="no2" then
    if int(emoney)=0 then
      notess="yes"
    elseif int(emoney)>0 and int(login_emoney)>=int(emoney) then
      conn.execute("update user_data set emoney=emoney-"&emoney&" where username='"&login_username&"'")
      conn.execute("insert into notes(username,nsort,iid,emoney,tim) values('"&login_username&"','"&n_sort&"',"&iid&","&emoney&",'"&now_time&"')")
      login_emoney=login_emoney-emoney:notess="yes"
    end if
  end if
  emoney_note=notess
end function
function web_var(wvar,wn)
  dim wdim,wnum:wnum=wn:wnum=wnum-1
  wdim=split(wvar,"|")
  if wnum>ubound(wdim) then web_var="":erase wdim:exit function
  web_var=wdim(wnum):erase wdim
end function
function web_varn(wvar,wn)
  dim wdim,wnum:wnum=wn:wnum=wnum-1
  wdim=split(wvar,"|")
  if wnum>ubound(wdim) then web_var=1:erase wdim:exit function
  web_varn=wdim(wnum):erase wdim
  if not(isnumeric(web_varn)) then web_varn=1
end function
function web_var_num(vvar,vnum,vn)
  if vnum>len(vvar) then web_var_num=0:exit function
  web_var_num=mid(vvar,vnum,vn)
  if not(isnumeric(web_var_num)) then web_var_num=0
end function
function var_null(ub)
  var_null=trim(ub)
  if var_null="" or isnull(var_null) then var_null=""
end function
function format_end(ft,fvar)
  if ft=0 then
    format_end="&nbsp;("&fvar&")"
  else
    format_end="&nbsp;<font class=gray>("&fvar&")</font>"
  end if
end function
function first_id(ndata)
  dim rsf
  set rsf=conn.execute("select top 1 id from "&ndata&" order by id desc")
  first_id=rsf("id")
  rsf.close:set rsf=nothing
end function
function format_user_view(uuser,ut,uc)
  if len(uuser)<1 then format_user_view="<font class=gray>-----</font>":exit function
  if uc<>"" then uc=" class="&uc
  format_user_view="<a href='user_view.asp?username="&server.urlencode(uuser)&"' title='查看 "&uuser&" 的详细资料'"
  if ut=1 then format_user_view=format_user_view&" target=_blank"
  format_user_view=format_user_view&uc&">"&uuser&"</a>"
end function
function format_img(fvar)
  format_img="<img border=0 src='images/"&web_var(web_config,5)&"/"&fvar&"'>"
end function
function icon_type(tn,tb)
  dim it_i
  for it_i=0 to tn
    icon_type=icon_type&"<img border=0 src='images/icon/"&it_i&".gif'> <input type=radio value="&it_i&" name=icon"
    if it_i=0 then icon_type=icon_type&" checked"
    icon_type=icon_type&" class=bg_"&tb&"> "
  next
end function
function img_small(snum)
  img_small="<img border=0 src='images/small/"&snum&".gif' align=absmiddle class=fr>&nbsp;"
end function
sub is_type()
%>
<%response.write img_small("isok")%>&nbsp;开放的主题&nbsp;
<%response.write img_small("ishot")%>&nbsp;回复超过10贴&nbsp;
<%response.write img_small("islock")%>&nbsp;锁定的主题&nbsp;
<%response.write img_small("istop")&"&nbsp;"&img_small("istops")%>&nbsp;固顶、总固顶的主题&nbsp;
<%response.write img_small("isgood")%>&nbsp;精华主题
<%
end sub
function left_action(jt,lat)
  dim jtn:jtn=img_small(jt)
  left_action=vbcrlf&"<table border=0 width='100%' cellspacing=0 cellpadding=4 align=center class=fr>" & _
		vbcrlf&"<tr><td height=5 width='50%'></td><td width='50%'></td></tr>" & _
		vbcrlf&"<tr><td>"&jtn&"<a href='user_action.asp?action=list'>用户列表</a></td><td>"&jtn&"<a href='online.asp'>与我在线</a></td></tr>" & _
		vbcrlf&"<tr><td>"&jtn&"<a href='user_action.asp?action=top'>发贴排行</a></td><td>"&jtn&"<a href='user_action.asp?action=emoney'>积分排行</a></td></tr>" & _
		vbcrlf&"<tr><td>"&jtn&"<a href='forum_action.asp?action=new'>论坛新贴</a></td><td>"&jtn&"<a href='forum_action.asp?action=hot'>热门话题</a></td></tr>" & _
		vbcrlf&"<tr><td>"&jtn&"<a href='forum_action.asp?action=top'>论坛置顶</a></td><td>"&jtn&"<a href='forum_action.asp?action=good'>论坛精华</a></td></tr>" & _
		vbcrlf&"<tr><td>"&jtn&"<a href='forum_action.asp?action=tim'>最新回复</a></td><td>"&jtn&"<a href='help.asp?action=forum'>论坛帮助</a></td></tr>" & _
		vbcrlf&"</table>"
  select case lat
  case 2
    left_action=kong&format_barc("<img src='images/"&web_var(web_config,5)&"/left_action.gif' border=0>",left_action,2,0,7)
  case 3
    left_action="<table border=0 width='100%' cellspacing=0 cellpadding=0 align=center><tr><td align=center>"&kong&format_bar("<img src='images/"&web_var(web_config,5)&"/left_action.gif' border=0>",left_action,0,2,jt,web_var(web_color,2)&"|"&web_var(web_color,6)&"|","")&"</td></tr></table>"
  case 4
    left_action="<table border=0 width='100%' cellspacing=0 cellpadding=0 align=center><tr><td align=center>"&kong&format_barc("<img src='images/"&web_var(web_config,5)&"/left_action.gif' border=0>",left_action,2,0,3)&"</td></tr></table>"
  case else
    left_action=kong&format_barc("<font class=end><b>功能跳转</b></font>",left_action,2,0,9)
  end select
end function

sub main_stat(sh,sjt,sm,st,sbg)
  dim num_topic,num_data,num_reg,new_username,num_news,num_article,num_down,num_flash,num_film,num_desktop,num_photo,stat_temp
  sql="select * from configs where id=1"
  set rs=conn.execute(sql)
  if rs.eof and rs.bof then
    rs.close
    conn.execute("insert into configs(id,num_topic,num_data,num_reg,new_username,num_news,num_article,num_down,num_product) values(1,0,0,0,'',0,0,0,0)")
    set rs=conn.execute(sql)
  end if
  num_topic=rs("num_topic")
  num_data=rs("num_data")
  num_reg=rs("num_reg")
  new_username=rs("new_username")
  num_news=rs("num_news")
  num_article=rs("num_article")
  num_down=rs("num_down")
  num_flash=rs("num_flash")
  num_film=rs("num_film")
  num_photo=rs("num_photo")
  num_desktop=rs("num_desktop")
  rs.close
  if sjt<>"" then sjt=img_small(sjt)
  stat_temp="<table border=0 width='100%' align=center><tr><td height=2></td></tr>"
  if st=1 then
    stat_temp=stat_temp&vbcrlf&"<tr><td>"&sjt&"网站版本：<font class=blue>"&web_var(web_stamp,int(mid(web_setup,3,1))+1)&"</font></td></tr>" & _
	      vbcrlf&"<tr><td>"&sjt&"新闻总数：<font class=red>"&num_news&"</font> 条</td></tr>" & _
	      vbcrlf&"<tr><td>"&sjt&"音乐总数：<font class=red>"&num_down&"</font> 个</td></tr>"& _
	      vbcrlf&"<tr><td>"&sjt&"视频总数：<font class=red>"&num_film&"</font> 个</td></tr>" & _
	      vbcrlf&"<tr><td>"&sjt&"Flash总数：<font class=red>"&num_flash&"</font> 个</td></tr>" & _
	      vbcrlf&"<tr><td>"&sjt&"照片总数：<font class=red>"&num_photo&"</font> 个</td></tr>" & _
	      vbcrlf&"<tr><td>"&sjt&"文章总数：<font class=red>"&num_article&"</font> 篇</td></tr>" & _
	      vbcrlf&"<tr><td>"&sjt&"壁纸总数：<font class=red>"&num_desktop&"</font> 张</td></tr>"
  end if
  stat_temp=stat_temp&vbcrlf&"<tr><td>"&sjt&"当前在线：<font class=red>"&online_num&"</font> 人</td></tr>" & _
            vbcrlf&"<tr><td>"&sjt&"网站注册：<font class=red>"&num_reg&"</font> 人</td></tr>" & _
	    vbcrlf&"<tr><td>"&sjt&"最新注册："&format_user_view(new_username,1,"")&"</td></tr>" & _
	    vbcrlf&"<tr><td>"&sjt&"主题总数：<font class=red>"&num_topic&"</font> 贴</td></tr>" & _
	    vbcrlf&"<tr><td>"&sjt&"贴子总数：<font class=red>"&num_data&"</font> 贴</td></tr>" & _
	    vbcrlf&"<tr><td height=2></td></tr></table>"
  if st=1 then
    call left_btype(stat_temp,"stat",sm,11)
  else
    response.write format_barc("<font class=end><b>数据统计</b></font>",stat_temp,2,0,5)
  end if
end sub
sub cookies_type(ct)
  response.cookies(web_cookies)("old_url")=request.servervariables("http_referer")
  response.cookies(web_cookies)("error_action")=ct
  call cookies_yes()
  call format_redirect("error.asp")
  response.end
end sub


sub cookies_yes()
  if request.cookies(web_cookies)("iscookies")="yes" then
    response.cookies(web_cookies).expires=date+365
  end if
end sub
sub format_redirect(fr)
  response.redirect fr
end sub


'*******************************************************************

'

'                     Beyondest.Com V3.6 Demo版

' 




'           网址：http://www.beyondest.com

' 

'*******************************************************************
%>
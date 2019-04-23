<%
' ====================
' Beyondest.Com v3.6.1
' http://beyondest.com
' ====================

function symbol_name(sn_var)
  dim safety_char:symbol_name="yes":sn_var=trim(sn_var)
  if sn_var="" or isnull(sn_var) or len(sn_var)>20 or instr(sn_var,"|")>0 or instr(sn_var,":")>0 or instr(sn_var,chr(9))>0 or instr(sn_var,chr(32))>0 or instr(sn_var,"'")>0 or instr(sn_var,"""")>0 then symbol_name="no":exit function
  safety_char=web_var(web_safety,1)
  for i=1 to len(sn_var)
    if instr(1,safety_char,mid(sn_var,i,1))>0 then symbol_name="no":exit function
  next
end function
function symbol_ok(symbol_var)
  dim safety_char:symbol_ok="yes":symbol_var=trim(symbol_var)
  if symbol_var="" or isnull(symbol_var) or len(symbol_var)>20 then symbol_ok="no":exit function
  safety_char=web_var(web_safety,2)
  for i=1 to len(symbol_var)
    if instr(1,safety_char,mid(symbol_var,i,1))=0 then symbol_ok="no":exit function
  next
end function
sub time_load(tt,t1,t2)
  if tt=0 then
    response.cookies(web_cookies)("time_load")=now_time
    call cookies_yes():exit sub
  end if
  dim tims,vars,tts
  tims=int(web_var(web_num,16))
  vars=time_type(request.cookies(web_cookies)("time_load"),9)
  if vars="" or not(isdate(vars)) then exit sub
  tts=int(DateDiff("s",vars,now_time))
  if tts<=tims then
    if t1=1 then set rs=nothing
    if t2=1 then call close_conn()
    call cookies_type("time_load")
    response.end
  end if
end sub
function health_name(hnn)
  health_name="yes":dim ti,tnum,tdim,hn
  hn=hnn:tdim=split(web_var(web_safety,3),":"):tnum=ubound(tdim)
  for ti=0 to tnum
    if instr(hn,tdim(ti))>0 then health_name="no":erase tdim:exit function
  next
  erase tdim
  tdim=split(web_var(web_safety,4),":"):tnum=ubound(tdim)
  for ti=0 to tnum
    if instr(hn,tdim(ti))>0 then health_name="no":erase tdim:exit function
  next
  erase tdim
end function
function health_var(hnn,vt)
  dim ti,tj,tdim,ht,hn:hn=hnn
  if int(mid(web_setup,4,1))=1 and vt=1 then
    tdim=split(web_var(web_safety,4),":")
    for ti=0 to ubound(tdim)
      ht=""
      for tj=1 to len(tdim(ti))
        ht=ht&"*"
      next
      hn=replace(hn,tdim(ti),ht)
    next
    erase tdim
  end if
  health_var=hn
end function
function found_error(error_type,error_len)
  if error_len>600 or error_len<200 then error_len=300
  found_error=VbCrLf & "<table border=0 width=" & error_len & " clas=fr><tr><td align=center height=50><font class=red>系统发现你输入的数据有以下错误：</font></td></tr><tr><td class=htd>" & error_type & "</td></tr><tr><td align=center height=50>" & go_back & "</td></tr></table>"
end function
function url_true(puu,pus)
  dim puuu,pu:puuu=puu:pu=pus
  if instr(1,pu,"://")<>0 then
    url_true=pu
  else
    if right(puuu,1)<>"/" then puuu=puuu&"/"
    url_true=puuu&pu
  end if
end function
Function time_type(tvar,tt)
  dim ttt:ttt=tvar
  if not(isdate(ttt)) then time_type="":exit function
  select case tt
  case 1	'10-10
    time_type=month(ttt)&"-"&day(ttt)
  case 11	'月-日
    time_type=month(ttt)&"月"&day(ttt)&"日"
  case 2	'年(2)-月-日 00-10-10
    time_type=right(year(ttt),2)&"-"&month(ttt)&"-"&day(ttt)
  case 3	'2000-10-10
    time_type=year(ttt)&"-"&month(ttt)&"-"&day(ttt)
  case 33	'年(4)-月-日
    time_type=year(ttt)&"年"&month(ttt)&"月"&day(ttt)&"日"
  case 4	'23:45
    time_type=hour(ttt)&":"&minute(ttt)
  case 44	'时:分
    time_type=hour(ttt)&"时"&minute(ttt)&"分"
  case 5	'23:45:36
    time_type=hour(ttt)&":"&minute(ttt)&":"&second(ttt)
  case 55	'时:分:秒
    time_type=hour(ttt)&"时"&minute(ttt)&"分"&second(ttt)&"秒"
  case 6	'10-10 23:45
    time_type=month(ttt)&"-"&day(ttt)&" "&hour(ttt)&":"&minute(ttt)
  case 66	'月-日 时:分
    time_type=month(ttt)&"月"&day(ttt)&"日 "&hour(ttt)&"时"&minute(ttt)&"分"
  case 7	'年(2)-月-日 时:分  00-10-10 23:45
    time_type=right(year(ttt),2)&"-"&month(ttt)&"-"&day(ttt)&" "&hour(ttt)&":"&minute(ttt)
  case 8	'2000-10-10 23:45
    time_type=year(ttt)&"-"&month(ttt)&"-"&day(ttt)&" "&hour(ttt)&":"&minute(ttt)
  case 88	'年(4)-月-日 时:分
    time_type=year(ttt)&"年"&month(ttt)&"月"&day(ttt)&"日 "&hour(ttt)&"时"&minute(ttt)&"分"
  case 9	'2000-10-10 23:45:45
    time_type=year(ttt)&"-"&month(ttt)&"-"&day(ttt)&" "&hour(ttt)&":"&minute(ttt)&":"&second(ttt)
    time_type=formatdatetime(time_type)
  case 99	'年(4)-月-日 时:分:秒
    time_type=year(ttt)&"年"&month(ttt)&"月"&day(ttt)&"日 "&hour(ttt)&"时"&minute(ttt)&"分"&second(ttt)&"秒"
  case else
    time_type=ttt
  end select
end function
function upload_time(tt)
  dim ttt:ttt=tt
  ttt=replace(ttt,":",""):ttt=replace(ttt,"-","")
  ttt=replace(ttt," ",""):ttt=replace(ttt,"/","")
  ttt=replace(ttt,"PM",""):ttt=replace(ttt,"AM","")
  ttt=replace(ttt,"上午",""):ttt=replace(ttt,"下午","")
  upload_time=ttt
end function
function code_form(strers)
  dim strer:strer=trim(strers)
  If isNull(strer) or strer="" Then code_form="":exit function
  strer=replace(strer,"'","""")
  code_form=strer
end function
function code_word(strers)
  dim strer:strer=trim(strers)
  If isNull(strer) or strer="" Then code_word="":exit function
  strer=replace(strer,"'","&#39;")
  code_word=strer
end function
function code_html(strers,chtype,cutenum)
  dim strer:strer=strers
  if isnull(strer) or strer="" then code_html="":exit function
  strer=health_var(strer,1)
  if cutenum>0 then strer=cuted(strer,cutenum)
  strer=replace(strer,"<","&lt;")
  strer=replace(strer,">","&gt;")
  strer=replace(strer,chr(39),"&#39;")		
  strer=replace(strer,chr(34),"&quot;")		
  strer=replace(strer,chr(32),"&nbsp;")		
  select case chtype
  case 1
    strer=replace(strer,chr(9),"&nbsp;")	
    strer=replace(strer,chr(10),"")		
    strer=replace(strer,chr(13),"")
  case 2
    strer=replace(strer,chr(9),"&nbsp;　&nbsp;")
    strer=replace(strer,chr(10),"<br>")		
    strer=replace(strer,chr(13),"<br>")
  end select
  code_html=strer
end function
function cuted(types,num)
  dim ctypes,cnum,ci,tt,tc,cc,cmod
  cmod=3
  ctypes=types:cnum=int(num):cuted="":tc=0:cc=0
  for ci=1 to len(ctypes)
    if cnum<0 then cuted=cuted&"...":exit for
    tt=mid(ctypes,ci,1)
    if int(asc(tt))>=0 then
      cuted=cuted&tt:tc=tc+1:cc=cc+1
      if tc=2 then tc=0:cnum=cnum-1
      if cc>cmod then cnum=cnum-1:cc=0
    else
      cnum=cnum-1
      if cnum<=0 then cuted=cuted&"...":exit for
      cuted=cuted&tt
    end if
  next
End Function
function post_chk()
  dim server_v1,server_v2
  post_chk="no"
  server_v1=Request.ServerVariables("HTTP_REFERER")
  server_v2=Request.ServerVariables("SERVER_NAME")
  if mid(server_v1,8,len(server_v2))=server_v2 then post_chk="yes":exit function
end function
function email_ok(email)
  dim names,name,i,c
  email_ok="yes":names = Split(email, "@")
  if UBound(names) <> 1 then email_ok="no":exit function
  for each name in names
    if Len(name) <= 0 then email_ok="no":exit function
    for i = 1 to Len(name)
      c = Lcase(Mid(name, i, 1))
      if InStr("abcdefghijklmnopqrstuvwxyz-_.", c) <= 0 and not IsNumeric(c) then email_ok="no":exit function
    next
    if Left(name, 1) = "." or Right(name, 1) = "." then email_ok="no":exit function
  next
  if InStr(names(1), ".") <= 0 then email_ok="no":exit function
  i = Len(names(1)) - InStrRev(names(1), ".")
  if i <> 2 and i <> 3 then email_ok="no":exit function
  if InStr(email, "..") > 0 then email_ok="no"
end function
function ip_sys(isu,iun)	'0,*=ip_sys  1,0=ip  1,1=ip:port  2,*=port  3,*=sys
  dim userip,userip2
  select case isu
  case 2
    ip_sys=Request.ServerVariables("REMOTE_PORT")
  case 3
    ip_sys=Request.Servervariables("HTTP_USER_AGENT")
  case else
    userip=Request.ServerVariables("HTTP_X_FORWARDED_FOR")
    userip2=Request.ServerVariables("REMOTE_ADDR")
    if instr(userip,",")>0 then userip=left(userip,instr(userip,",")-1)
    if instr(userip2,",")>0 then userip2=left(userip2,instr(userip2,",")-1)
    if userip="" then
      ip_sys=userip2
    else
      ip_sys=userip
    end if
    if isu=0 then ip_sys="您的真实ＩＰ是：" & ip_sys & ":" & Request.ServerVariables("REMOTE_PORT") & "，" & view_sys(Request.Servervariables("HTTP_USER_AGENT")):exit function
    if iun=1 then ip_sys=ip_sys&":"&Request.ServerVariables("REMOTE_PORT")
  end select
end function
function view_sys(vss)
  dim vs,ver,strUserAgentArr,strTempUserInfo1,strTempUserInfo2,Mozilla,Agent,BcType,Browser,sSystem,strSystem,strBrowser
  on error resume next
  vs=trim(vss):strUserAgentArr=Split(vs, "; "):strTempUserInfo1=strUserAgentArr(1)
  if Instr(strTempUserInfo1, "MSIE") > 0 then strTempUserInfo1=Replace(strTempUserInfo1, "MSIE", "Internet Explorer")
  strTempUserInfo2=trim(Left(strUserAgentArr(2), Len(strUserAgentArr(2))-1))
  if InStr(vs, "Mozilla/4.0 (compatible;") > 0 and strTempUserInfo2="Windows NT 5.0" then strTempUserInfo2="Windows 2000"
  Mozilla=vs:Agent=Split(Mozilla,"; "):BcType=0
  If Instr(Agent(1),"U") Or Instr(Agent(1),"I") Then BcType=1
  If InStr(Agent(1),"MSIE") Then BcType=2
  Select Case BcType
  Case 0
    Browser="其它":sSystem="其它"
  Case 1
    Ver=Mid(Agent(0),InStr(Agent(0), "/")+1)
    Ver=Mid(Ver,1,InStr(Ver, " ")-1)
    Browser="Netscape" & Ver
    sSystem=Mid(Agent(0), InStr(Agent(0), "(")+1)
    sSystem=Replace(sSystem, "Windows", "Win")
  case 2
    Browser=Agent(1):sSystem=Replace(Agent(2), ")", ""):sSystem=Replace(sSystem, "Windows", "Win")
  End Select
  strSystem=Replace(sSystem, "Win", "Windows")
  if InStr(strSystem,"98")>0 and InStr(Mozilla,"Win 9x")>0 then strSystem=Replace(strSystem, "98", "Me")
  strSystem=Replace(strSystem, "NT 5.0", "2000")
  strSystem=Replace(strSystem, "NT5.0", "2000")
  strSystem=Replace(strSystem, "NT 5.1", "XP")
  strSystem=Replace(strSystem, "NT5.1", "XP")
  strSystem=Replace(strSystem, "NT 5.2", "2003")
  strSystem=Replace(strSystem, "NT5.2", "2003")
  strBrowser=Replace(Browser, "MSIE", "Internet Explorer")
  set Browser=Nothing:set sSystem=Nothing
  view_sys="操作系统：" & trim(strSystem) & " ，浏览器：" & trim(strBrowser)
  if err then err.clear:view_sys="未知的操作系统和浏览器"
end function
function ip_true(tips)
  dim tip,iptemp,iptemp1,iptemp2:tip=tips:ip_true="no":tip=trim(tip)
  iptemp=tip:iptemp=replace(replace(iptemp,".",""),":","")
  iptemp1=tip:iptemp1=len(tip)-len(replace(iptemp1,".",""))
  iptemp2=tip:iptemp2=len(tip)-len(replace(iptemp2,":",""))
  if isnumeric(iptemp) and cint(iptemp1)=3 and (cint(iptemp2)=1 or cint(iptemp2)=0) then ip_true="yes"
end function
function ip_ip(tips)
  dim ipn,tip:tip=tips:tip=trim(tip)
  if ip_true(tip)="no" then ip_ip="no":exit function
  ipn=Instr(tip,":")
  if ipn>0 then ip_ip=left(tip,ipn-1):exit function
  ip_ip=tip
end function
function ip_types(tips,tu,tt)
  dim ipn,tip2,wip,ip_type:tip2=tips:tip2=trim(tip2):ip_types="error"
  if ip_true(tip2)="no" then ip_types="no":exit function
  wip=int(mid(web_setup,5,1))
  if login_mode=format_power2(1,1) then wip=2
  select case wip
  case 0
    ip_types="*.*.*.*":ip_types=tu&" 的IP是："&ip_types
    if tt<>0 then ip_types="<img src='images/small/ip.gif' align=absMiddle title='"&ip_types&"' border=0>"
  case 1
    ipn=Instr(tip2,":")
    if ipn>0 then
      ip_types=left(tip2,ipn-1)
    else
      ip_types=tip2
    end if
    ip_type=split(ip_types,"."):ip_types=ip_type(0)&"."&ip_type(1)&".*.*"
    erase ip_type:ip_types=tu&" 的IP是："&ip_types
    if tt<>0 then ip_types="<img src='images/small/ip.gif' align=absMiddle title='"&ip_types&"' border=0>"
  case else
    ip_types=tu&" 的IP是："&tip2
    if tt<>0 then ip_types="<a href='ip_address.asp?ip="&tip2&"'><img src='images/small/ip.gif' align=absMiddle title='"&ip_types&"' border=0></a>"
  end select
end function
function ip_port(pips)
  dim ipnn,iptemp,pip:pip=pips
  pip=trim(pip)
  if ip_true(pip)="no" then ip_port="no":exit function
  ipnn=Instr(pip,":")
  if ipnn>0 then ip_port=right(pip,len(pip)-ipnn):exit function
  ip_port="yes"
end function
function ip_address(sips)
  dim str1,str2,str3,str4,num,country,city,irs,sip:sip=sips
  if not(isnumeric(left(sip,2))) then ip_address="未知":exit function
  if sip="127.0.0.1" then sip="192.168.0.1"
  str1=left(sip,instr(sip,".")-1):sip=mid(sip,instr(sip,".")+1)
  str2=left(sip,instr(sip,".")-1):sip=mid(sip,instr(sip,".")+1)
  str3=left(sip,instr(sip,".")-1):str4=mid(sip,instr(sip,".")+1)
  if not(isNumeric(str1)=0 or isNumeric(str2)=0 or isNumeric(str3)=0 or isNumeric(str4)=0) then
    num=cint(str1)*256*256*256+cint(str2)*256*256+cint(str3)*256+cint(str4)-1
    sql="select Top 1 country,city from ip_address where ip1 <="&num&" and ip2 >="&num&""
    set irs=server.createobject("adodb.recordset")
    irs.open sql,conn,1,1
    if irs.eof and irs.bof then 
      country="亚洲":city=""
    else
      country=irs(0):city=irs(1)
    end if
    irs.close:set irs=nothing
  end if
  ip_address=country&city
end function
%>
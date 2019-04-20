<%
'*******************************************************************
'
'                     Beyondest.Com V3.6 Demo版
' 
'           网址：http://www.beyondest.com
' 
'*******************************************************************

sub configs_load()
  conn.execute("insert into configs(id,num_topic,num_data,num_reg,new_username,num_news,num_article,num_down,num_product,counter,max_online,max_tim,start_tim) values(1,0,0,0,'',0,0,0,0,1,1,'"&now_time&"','"&time_type(now_time,33)&"')")
end sub

function counter_type(cmain,ct)
  dim rs,sql,ft,counters,max_online,max_tim,start_tim,types,counts
  types=1:i=0:counter_type=""
  sql="select counter,max_online,max_tim,start_tim from configs where id=1"
  set rs=conn.execute(sql)
  if rs.eof and rs.bof then
    rs.close
    call configs_load()
    set rs=conn.execute(sql)
  end if
  counters=rs("counter")
  max_online=rs("max_online")
  max_tim=rs("max_tim")
  start_tim=rs("start_tim")
  rs.close:set rs=nothing
  counters=int(counters)
  max_online=int(max_online)
  ft=mid(web_setup,7,1)
  if not(isnumeric(ft)) then ft=1
  ft=int(ft)
  if ft=0 then
    if trim(request.cookies(web_cookies)("counters"))<>"yes" then
      counters=counters+types
      response.cookies(web_cookies)("counters")="yes"
    end if
  else
    response.cookies(web_cookies)("counters")=""
    counters=counters+types
  end if
  if online_num>max_online then max_online=online_num:max_tim=now_time
  sql="update configs set counter="&counters&",max_online="&max_online&",max_tim='"&max_tim&"' where id=1"
  conn.execute(sql)
  if cmain="view" then
    counts="本站总访问量:&nbsp;<font class=red_3 title=从&nbsp;"&start_tim&"&nbsp;至今>" & counters & "</font>&nbsp;人次" & _
	   "&nbsp;┋&nbsp;最高峰&nbsp;<font class=red_3 title=最高峰发生在：" & max_tim & ">" & max_online & "</font>&nbsp;人在线" & _
	   "&nbsp;┋&nbsp;当前有&nbsp;<font class=red_3>" & online_num & "</font>&nbsp;人在线"
    counter_type=counts
  end if
end function
%>
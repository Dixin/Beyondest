<!-- #include file="include/onlogin.asp" -->
<%
'*******************************************************************
'
'                     Beyondest.Com V3.6 Demo版
' 
'           网址：http://www.beyondest.com
' 
'*******************************************************************

dim classid,forumid
classid=trim(request("class_id"))
forumid=trim(request("forum_id"))
action=trim(request("action"))
if not(isnumeric(classid)) or not(isnumeric(forumid)) or (action<>"up" and action<>"down") then
  response.redirect "admin_forum.asp"
  response.end
end if
%>
<!-- #include file="include/conn.asp" -->
<%
dim tmp_id_1,tmp_id_2,tmp_order_1,tmp_order_2,sqladd,update_ok,rssum
update_ok="no"
if action="up" then
  sqladd=" desc"
else
  sqladd=""
end if

sql="select forum_order from bbs_forum where forum_id="&forumid&" and class_id="&classid
set rs=conn.execute(sql)
if rs.eof and rs.bof then
  rs.close:set rs=nothing
  close_conn
  response.redirect "admin_forum.asp"
  response.end
end if
rs.close:set rs=nothing

sql="select forum_id,forum_order from bbs_forum where class_id="&classid&" order by forum_order"&sqladd&",forum_id desc"
set rs=conn.execute(sql)
do while not rs.eof
  if int(rs("forum_id"))=int(forumid) then
    tmp_id_1=forumid
    tmp_order_1=rs("forum_order")
    rs.movenext
    if not rs.eof then
      tmp_id_2=rs("forum_id")
      tmp_order_2=rs("forum_order")
      update_ok="yes"
      exit do
    end if
    exit do
  end if
  rs.movenext
loop
rs.close:set rs=nothing

if update_ok="yes" then
  sql="update bbs_forum set forum_order="&tmp_order_2&" where forum_id="&tmp_id_1
  response.write sql
  conn.execute(sql)
  sql="update bbs_forum set forum_order="&tmp_order_1&" where forum_id="&tmp_id_2
  conn.execute(sql)
end if

close_conn

response.redirect "admin_forum.asp"
%>
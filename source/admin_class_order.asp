<!-- #include file="include/onlogin.asp" -->
<%
' ====================
' Beyondest.Com v3.6.1
' http://beyondest.com
' ====================

dim classid
classid=trim(request("class_id"))
action=trim(request("action"))
if not(isnumeric(classid)) or (action<>"up" and action<>"down") then
  response.redirect "admin_forum.asp"
  response.end
end if
%>
<!-- #include file="include/conn.asp" -->
<%
dim tmp_id_1,tmp_id_2,tmp_order_1,tmp_order_2,sqladd,update_ok
update_ok="no"
if action="up" then
  sqladd=" desc"
else
  sqladd=""
end if

sql="select * from bbs_class where class_id="&classid
set rs=conn.execute(sql)
if rs.eof and rs.bof then
  rs.close:set rs=nothing
  close_conn
  response.redirect "admin_forum.asp"
  response.end
end if
rs.close:set rs=nothing

sql="select * from bbs_class order by class_order"&sqladd
set rs=conn.execute(sql)
do while not rs.eof
  if int(rs("class_id"))=int(classid) then
    tmp_id_1=classid
    tmp_order_1=rs("class_order")
    rs.movenext
    if not rs.eof then
      tmp_id_2=rs("class_id")
      tmp_order_2=rs("class_order")
      update_ok="yes"
      exit do
    end if
    exit do
  end if
  rs.movenext
loop
rs.close:set rs=nothing

if update_ok="yes" then
  sql="update bbs_class set class_order="&tmp_order_2&" where class_id="&tmp_id_1
  conn.execute(sql)
  sql="update bbs_class set class_order="&tmp_order_1&" where class_id="&tmp_id_2
  conn.execute(sql)
end if

close_conn

response.redirect "admin_forum.asp"
%>
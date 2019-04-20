<!-- #include file="config.asp" -->
<!-- #include file="config_frm.asp" -->
<!-- #include file="config_upload.asp" -->
<!-- #include file="config_put.asp" -->
<%
'*******************************************************************
'
'                     Beyondest.Com V3.6 Demo版
' 
'           网址：http://www.beyondest.com
' 
'*******************************************************************

if session("beyondest_online_admin")<>"beyondest_admin" then
  response.redirect "admin_login.asp"
  response.end
end if
if web_login=0 then web_login=2
dim color1,color2,color3,table1,mtr
color1=web_var(web_color,1)
color2=web_var(web_color,5)
color3=web_var(web_color,6)
table1=" bordercolorlight=#c0c0c0 bordercolordark="&color1
mtr=" onmouseover=""javascript:this.bgColor='"&color2&"';"" onmouseout=""javascript:this.bgColor='"&color1&"';"""
function del_select(delid)
  dim del_i,del_num,del_dim,del_sql,del_rs,del_username,fobj,picc
  if delid<>"" and not isnull(delid) then
    delid=replace(delid," ","")
    del_dim=split(delid,",")
    del_num=UBound(del_dim)
    for del_i=0 to del_num
      'del_sql
      del_sql="select username from "&data_name&" where id="&del_dim(del_i)
      set del_rs=conn.execute(del_sql)
      if not(del_rs.eof and del_rs.bof) then
        call user_integral("del",web_varn(web_num,15),del_rs("username"))
      end if
      del_rs.close:set del_rs=nothing
      call upload_del(data_name,del_dim(del_i))
      del_sql="delete from "&data_name&" where id="&del_dim(del_i)
      conn.execute(del_sql)
    next
    Erase del_dim
    del_select=vbcrlf&"<script language=javascript>alert(""共删除了 "&del_num+1&" 条记录！"");</script>"
  end if
end function
function header(popedomnum,titmenu)
  if session("beyondest_online_admines")<>web_var(web_config,3) then
    if session("beyondest_online_admines")<>"beyondest" and popedom_formated(session("beyondest_online_popedom"),popedomnum,0)=0 then
      response.redirect "admin.asp?action=main&error=popedom"
      response.end
    end if
  end if
  header = VbCrLf & "<html><head><title>"&web_var(web_config,1)&" - 管理后台</title>" & _
	   VbCrLf & "<meta http-equiv=Content-Type content=text/html; charset=gb2312>" & _
	   VbCrLf & "<link rel=stylesheet href='include/beyondest.css' type=text/css>" & _
	   VbCrLf & "<script langiage='javascript' src='style/open_win.js'></script>" & _
	   VbCrLf & "<script langiage='javascript' src='style/mouse_on_title.js'></script>" & _
	   VbCrLf & "</head>" & VbCrLf & "<body topmargin=0 leftmargin=0 bgcolor="&color1&"><center>" & _
	   VbCrLf & "<table border=0 width=600 cellspacing=0 cellpadding=0>" & _
	   vbcrlf & "<tr><td height=50 align=center>"&titmenu&"&nbsp;┋&nbsp;<a href='javascript:;' onclick=""javascript:document.location.reload()"">刷新</a></td></tr><tr><td align=center height=350>"
end function
function popedom_formated(popedom1,popedomnum,popedomtype)
  dim poptemp:poptemp=0
  if len(popedom1)=50 and popedomnum<>-1 then
    poptemp=mid(popedom1,popedomnum,1)
  end if
  if popedomtype<>0 then
    if poptemp=0 then
      poptemp=1
    else
      poptemp=0
    end if
  end if
  if poptemp<>0 then poptemp=1
  if popedomnum=-1 then poptemp=1
  popedom_formated=poptemp
end function
%>
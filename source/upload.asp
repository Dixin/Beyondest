<!-- #INCLUDE file="include/config.asp" -->
<!--#include file="INCLUDE/upload_config.asp"-->
<!--#include file="include/conn.asp"-->
<%
' ====================
' Beyondest.Com v3.6.1
' http://beyondest.com
' ====================
%>
<script langiage='javascript' src='style/beyondest_config.js'></script>
<script langiage='javascript' src='STYLE/mouse_on_title.js'></script>
<html>
<head>
<title><%response.write web_var(web_config,1)%> - 文件上传</title>
<meta http-equiv=Content-Type content=text/html; charset=gb2312>
<link rel=stylesheet href='include/beyondest.css' type=text/css>
</head>
<body topmargin=0 leftmargin=0 bgcolor=<%response.write web_var(web_color,1)%>>
<table border=0 height='100%'cellspacing=0 cellpadding=0><tr><td height='100%'>
<%
dim formname,upload_path,upload_type,upload_size,uup
uup="|article|down|forum|gallery|news|other|product|video|website|"
action=trim(request.querystring("action"))
if var_null(login_username)="" or var_null(login_password)="" then
  response.write web_var(web_error,2)
elseif post_chk()="no" then
  response.write web_var(web_error,1)
else
  upload_path=web_var(web_upload,1)
  upload_type=web_var(web_upload,2)
  upload_size=web_var(web_upload,3)
  select case action
  case "upfile"
    call upload_chk()
  case else
    call upload_main()
  end select
end if

call close_conn()

sub upload_chk()
  Server.ScriptTimeOut=5000
  dim upload,up_text,up_path,uptemp,uppath,up_name,MyFso,upfile,upcount,upfilename,upfile_name,upfile_name2,upfilesize,upid
  set upload=new upload_classes
  set MyFso=CreateObject("Scripting.FileSystemObject")
  upcount=1
  up_name=trim(upload.form("up_name"))
  up_text=trim(upload.form("up_text"))
  up_path=trim(upload.form("up_path"))
  if session("beyondest_online_admin")<>"beyondest_admin" and len(up_name)>2 then up_name=""
  if len(up_name)<3 then up_name=up_name&upload_time(now_time)
  if int(instr(uup,"|"&up_path&"|"))=0 then up_path="other"
  if len(up_path)<3 then up_path="other"
  uppath=up_path
  if right(upload_path,1)<>"/" then upload_path=upload_path&"/"
  up_path=server.mappath(upload_path&up_path)
  if not MyFso.folderExists(up_path) then
    set up_path=MyFso.CreateFolder(up_path)
  end if
  if right(up_path,1)<>"/" then up_path=up_path&"/"
  'if right(uppath,1)<>"/" then uppath=uppath&"/"

  set upfile=upload.file("file_name1")
  upfilesize=upfile.FileSize
  upload_size=upload_size*1024
  if upfilesize>0 then
    upfilename=upfile.FileName
    if upfilesize>upload_size then
      uptemp="<font class=red_2>上传失败</font>：文件太大！(不能超过"&int(upload_size/1024)&"KB) "&go_back
    else
      upfile_name=Right(upfilename,(len(upfilename)-Instr(upfilename,".")))
      upfile_name=lcase(upfile_name)
      if instr(","&upload_type&",",","&upfile_name&",")>0 then
        upfile_name2=upfile_name
        upfile_name=up_name&"."&upfile_name
        upfile.SaveAs up_path&upfile_name
        
        sql="select id from upload where url='"&uppath&"/"&upfile_name&"'"
        set rs=conn.execute(sql)
        if rs.eof and rs.bof then
          rs.close
          sql="insert into upload(iid,nsort,types,username,url,genre,sizes,tim) " & _
	      "values(0,'',0,'"&login_username&"','"&uppath&"/"&upfile_name&"','"&upfile_name2&"',"&upfilesize&",'"&now_time&"')"
          conn.execute(sql)
          sql="select top 1 id from upload order by id desc"
          set rs=conn.execute(sql)
          upid=int(rs("id"))
        else
          conn.execute("update upload set username='"&login_username&"',sizes="&upfilesize&",tim='"&now_time&"' where id="&rs("id"))
        end if
        rs.close:set rs=nothing
        
        uptemp="<font class=red>上传成功</font>：<a href='"&upload_path&uppath&"/"&upfile_name&"' target=_blank>"&upfile_name&"</a> ("&upfilesize&"Byte)"
        if instr(up_text,"pic")>0 then
          uptemp=uptemp&"<script>parent.document.all."&up_text&".value='"&uppath&"/"&upfile_name&"';"
        else
          uptemp=uptemp&"&nbsp;&nbsp;[ <a href='?uppath="&uppath&"&upname=&uptext="&up_text&"'>点击继续上传</a> ]<script>parent.document.all."&up_text&".value+='"
          select case lcase(upfile_name2)
          case "gif","jpg","bmp","png"
            uptemp=uptemp&"[IMG]"&upload_path&uppath&"/"&upfile_name&"[/IMG]"
          case "swf"
            uptemp=uptemp&"[FLASH="&web_var(web_num,9)&","&web_var(web_num,10)&"]"&upload_path&uppath&"/"&upfile_name&"[/FLASH]"
          case else
            uptemp=uptemp&"[DOWNLOAD]"&upload_path&uppath&"/"&upfile_name&"[/DOWNLOAD]"
          end select
          uptemp=uptemp&"\n';"
        end if
        if int(upid)>0 then uptemp=uptemp&"parent.document.all.upid.value+=',"&upid&"';"
        uptemp=uptemp&"</script>"
      else
        uptemp="<font class=red_2>上传失败</font>：文件类型只能为："&replace(upload_type,"|","、")&"等格式) "&go_back
      end if
    end if
  else
    uptemp="<font class=red_2>上传失败</font>：您可能没有选择想要上传的文件！"&go_back
  end if
  set upfile=nothing
  response.write uptemp
  set MyFso=nothing
  set upload=nothing
end sub

sub upload_main()
  dim uppath,upname,uptext
  uppath=trim(request.querystring("uppath"))
  upname=trim(request.querystring("upname"))
  uptext=trim(request.querystring("uptext"))
  if session("beyondest_online_admin")<>"beyondest_admin" and len(upname)>2 then upname=""
  if int(instr(uup,"|"&uppath&"|"))=0 then response.write "参数出错！":exit sub
  if len(uppath)<1 or len(uptext)<1 then response.write "参数出错！":exit sub
  'if len(upname)<3 then upname=upname&upload_time(now_time)
%>
<table border=0 cellspacing=0 cellpadding=2>
<form name=form1 action='?action=upfile' method=post enctype='multipart/form-data'>
<input type=hidden name=up_path value='<%response.write uppath%>'>
<input type=hidden name=up_name value='<% response.write upname %>'>
<input type=hidden name=up_text value='<% response.write uptext %>'>
<tr>
<td><input type=file name=file_name1 value='' size=35></td>
<td align=center height=30><input type=submit name=submit value='点击上传'> (<=<% response.write upload_size %>KB)</td>
</tr>
</form>
</table>
<% end sub %>
</td></tr></table>
</body>
</html>

<%
' ====================
' Beyondest.Com v3.6.1
' http://beyondest.com
' ====================

sub del_file(fn)
  on error resume next
  dim fobj,picc,upload_path
  picc=fn:upload_path=web_var(web_upload,1)
  if len(picc)>3 then
    if int(instr(1,picc,"://"))=0 then
      if left(picc,1)<>"/" then
        if right(upload_path,1)<>"/" then upload_path=upload_path&"/"
        picc=Server.MapPath(upload_path&picc)
      else
        picc=Server.MapPath(picc)
      end if
      Set fobj = CreateObject("Scripting.FileSystemObject")
      fobj.DeleteFile(picc)
      Set fobj = nothing
    end if
  end if
end sub

sub upload_del(nsort,iid)
  dim rs,sql
  sql="select url from upload where nsort='"&nsort&"' and iid="&iid&" order by id"
  set rs=conn.execute(sql)
  do while not rs.eof
    call del_file(rs("url"))
    rs.movenext
  loop
  rs.close:set rs=nothing
  conn.execute("delete from upload where nsort='"&nsort&"' and iid="&iid)
end sub

sub upload_note(ns,iid)
  dim ddim,i,sql,upid:upid=trim(request.form("upid"))
  if len(upid)<1 then exit sub
  if left(upid,1)="," then upid=right(upid,len(upid)-1)
  ddim=split(upid,",")
  for i=0 to ubound(ddim)
    conn.execute("update upload set iid="&iid&",nsort='"&ns&"',types=1 where id="&ddim(i))
  next
  erase ddim
end sub
%>
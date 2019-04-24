<%
' ====================
' Beyondest.Com v3.6.1
' http://beyondest.com
' ====================
Sub del_file(fn)
    On Error Resume Next
    Dim fobj,picc,upload_path
    picc = fn:upload_path = web_var(web_upload,1)

    If Len(picc) > 3 Then

        If Int(InStr(1,picc,"://")) = 0 Then

            If Left(picc,1) <> "/" Then
                If Right(upload_path,1) <> "/" Then upload_path = upload_path & "/"
                picc = Server.MapPath(upload_path & picc)
            Else
                picc = Server.MapPath(picc)
            End If

            Set fobj = CreateObject("Scripting.FileSystemObject")
            fobj.DeleteFile(picc)
            Set fobj = Nothing
        End If

    End If

End Sub

Sub upload_del(nsort,iid)
    Dim rs,sql
    sql    = "select url from upload where nsort='" & nsort & "' and iid=" & iid & " order by id"
    Set rs = conn.execute(sql)

    Do While Not rs.eof
        Call del_file(rs("url"))
        rs.movenext
    Loop

    rs.Close:Set rs = Nothing
    conn.execute("delete from upload where nsort='" & nsort & "' and iid=" & iid)
End Sub

Sub upload_note(ns,iid)
    Dim ddim,i,sql,upid:upid = Trim(Request.form("upid"))

    If Len(upid) < 1 Then Exit Sub
        If Left(upid,1) = "," Then upid = Right(upid,Len(upid) - 1)
        ddim = Split(upid,",")

        For i = 0 To UBound(ddim)
            conn.execute("update upload set iid=" & iid & ",nsort='" & ns & "',types=1 where id=" & ddim(i))
        Next

        Erase ddim
    End Sub %>
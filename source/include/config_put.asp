<%
' ====================
' Beyondest.Com v3.6.1
' http://beyondest.com
' ====================
Function code_admin(strers)
    Dim strer:strer = Trim(strers)
    If IsNull(strer) Or strer = "" Then code_admin = "":Exit Function
    strer      = Replace(strer,"'","""")
    code_admin = strer
End Function

Function code_config(strers)
    Dim strer:strer = Trim(strers)
    If IsNull(strer) Or strer = "" Then code_config = "":Exit Function
    strer       = Replace(strer,"'","")
    strer       = Replace(strer,":","")
    strer       = Replace(strer,"|","")
    strer       = Replace(strer,".","")
    code_config = strer
End Function

Function ender()
    ender = vbcrlf & "</td></tr><tr><td height=30 align=center>" & web_var(web_error,4) & "</td></tr>" & _
    vbcrlf & "</center></body></html>"
End Function

Sub sql_cid_sid()

    If cid <> 0 Then
        sqladd      = " where c_id=" & cid
        pageurl     = pageurl & "c_id=" & cid & "&"

        If sid <> 0 Then
            sqladd  = sqladd & " and s_id=" & sid
            pageurl = pageurl & "s_id=" & sid & "&"
        End If

    End If

End Sub

Sub chk_cid_sid()

    If IsNumeric(csid) Then
        cid = csid:sid = 0
    Else
        cid = Mid(csid,1,InStr(csid,"-") - 1)
        sid = Mid(csid,InStr(csid,"-") + 1,Len(csid))
    End If

End Sub

Sub admin_cid_sid()
    cid = Trim(Request.querystring("c_id"))
    sid = Trim(Request.querystring("s_id"))
    If Not(IsNumeric(cid)) Then cid = 0
    If Not(IsNumeric(sid)) Then sid = 0
    cid = Int(cid):sid = Int(sid)
End Sub

Sub chk_csid(cid,sid)
    Dim sql3,rs3
    Response.Write "<select name=csid size=1>"
    sql3    = "select c_id,c_name from jk_class where nsort='" & nsort & "' order by c_order,c_id"
    Set rs3 = conn.execute(sql3)

    Do While Not rs3.eof
        nid = Int(rs3(0))
        Response.Write vbcrlf & "<option value='" & nid & "' class=bg_2"
        If cid = nid Then Response.Write " selected"
        Response.Write ">" & rs3(1) & "</option>"
        sql2       = "select s_id,s_name from jk_sort where c_id=" & nid & " order by s_order,s_id"
        Set rs2    = conn.execute(sql2)

        Do While Not rs2.eof
            now_id = Int(rs2(0))
            Response.Write vbcrlf & "<option value='" & nid & "-" & now_id & "'"
            If sid = now_id Then Response.Write " selected"
            Response.Write ">　" & rs2(1) & "</option>"
            rs2.movenext
        Loop

        rs2.Close:Set rs2 = Nothing
        rs3.movenext
    Loop

    rs3.Close:Set rs3 = Nothing
    Response.Write "</select>" & redx
End Sub

Sub left_sort()
    Dim rs,sql
    sql     = "select c_id,c_name from jk_class where nsort='" & nsort & "' order by c_order,c_id"
    Set rs  = conn.execute(sql)

    Do While Not rs.eof
        nid = Int(rs(0))

        If cid = nid Then
            Response.Write vbcrlf & img_small("jt1") & "<a href='?c_id=" & nid & "'><b><font class=red_3>" & rs(1) & "</b></font></a><br>"
        Else
            Response.Write vbcrlf & img_small("jt0") & "<a href='?c_id=" & nid & "'><font class=red_3>" & rs(1) & "</font></a><br>"
        End If

        sql2       = "select s_id,s_name from jk_sort where c_id=" & nid & " order by s_order,s_id"
        Set rs2    = conn.execute(sql2)

        Do While Not rs2.eof
            now_id = Int(rs2(0))

            If sid = now_id Then
                Response.Write vbcrlf & "　<a href='?c_id=" & nid & "&s_id=" & now_id & "'><font class=blue>" & rs2(1) & "</a></a><br>"
            Else
                Response.Write vbcrlf & "　<a href='?c_id=" & nid & "&s_id=" & now_id & "'>" & rs2(1) & "</a><br>"
            End If

            rs2.movenext
        Loop

        rs2.Close:Set rs2 = Nothing
        rs.movenext
    Loop

    rs.Close
End Sub

Sub left_sort2()
    Dim rs,sql
    sql     = "select c_id,c_name from jk_class where nsort='" & nsort & "' order by c_order,c_id"
    Set rs  = conn.execute(sql)

    Do While Not rs.eof
        nid = Int(rs(0))

        If cid = nid Then
            Response.Write vbcrlf & img_small("jt1") & "<a href='?types=" & types & "&c_id=" & nid & "'><b><font class=red_3>" & rs(1) & "</b></font></a><br>"
        Else
            Response.Write vbcrlf & img_small("jt0") & "<a href='?types=" & types & "&c_id=" & nid & "'><font class=red_3>" & rs(1) & "</font></a><br>"
        End If

        sql2 = "select s_id,s_name from jk_sort where c_id=" & nid & " order by s_order,s_id"
        Set rs2 = conn.execute(sql2)

        Do While Not rs2.eof
            now_id = Int(rs2(0))

            If sid = now_id Then
                Response.Write vbcrlf & "　<a href='?types=" & types & "&c_id=" & nid & "&s_id=" & now_id & "'><font class=blue>" & rs2(1) & "</a></a><br>"
            Else
                Response.Write vbcrlf & "　<a href='?types=" & types & "&c_id=" & nid & "&s_id=" & now_id & "'>" & rs2(1) & "</a><br>"
            End If

            rs2.movenext
        Loop

        rs2.Close:Set rs2 = Nothing
        rs.movenext
    Loop

    rs.Close
End Sub

Sub chk_power(power,pt)
    Dim ddim:ddim = Split(user_power,"|")

    For i = 0 To UBound(ddim)
        Response.Write vbcrlf & "<input type=checkbox name=power value='" & i + 1 & "' class=bg_1"
        If InStr(1,"." & power & ".","." & i + 1 & ".") > 0 Or pt = 1 Then Response.Write " checked"
        Response.Write ">" & Right(ddim(i),Len(ddim(i)) - InStr(ddim(i),":"))
    Next

    Erase ddim %><input type=checkbox name=power value='0' class=bg_1<% If InStr(1,"." & power & ".",".0.") > 0 Then Response.Write " checked" %>>游客<%
End Sub

Sub chk_emoney(ee)
    Response.Write "&nbsp;货币：<input type=text name=emoney value='" & ee & "' size=6 maxlength=10>"
End Sub

Sub chk_h_u() %>&nbsp;&nbsp;<input type=checkbox name=hidden<% If rs("hidden") = False Then Response.Write " checked" %> value='yes'>&nbsp;隐藏
&nbsp;<input type=checkbox name=username_my value='yes'>&nbsp;<font alt='发布人：<% Response.Write rs("username") %>'>修改发布人为我</font><%
End Sub %>
<% @LANGUAGE = "VBSCRIPT" CODEPAGE = "936" %>
<%
Option Explicit
'Response.buffer=true 		'开启缓冲页面功能，只要把前面的引号去掉就可以开启了

Dim web_config,web_cookies,web_login,web_setup,web_menu,web_num,web_color,web_upload,web_safety,web_edition,web_label,web_lct,web_imglct
Dim web_news_art,web_down,web_shop,web_stamp,space_mod,gang,gang2,sk_bar_lf
Dim web_error,user_grade,user_power,forum_type,now_time,timer_start,redx,kong,ukong,go_back,closer,lefter,righter
Dim login_username,login_password,login_mode,login_message,login_popedom,login_emoney
Dim rs,sql,i,tit,tit_fir,index_url,action,page_power,m_unit,sk_bar,sk_class,sk_jt,sk_img %>
<!-- #include file="common.asp" -->
<!-- #include file="functions.asp" -->
<%
sk_bar         = 11:sk_bar_lf = 15:sk_class = "end":sk_jt = "jt0":space_mod = web_var(web_num,12):m_unit = web_var(web_config,8):now_time = time_type(Now(),9):timer_start = timer()
web_label      = "<a href='http://beyondest.com/' target=_blank>Power by <b><font face=Arial color=#CC3300>Beyondest</font><font face=Arial>.Com</font></b></a>"
redx           = "&nbsp;<font color='#ff0000'>*</font>&nbsp;":kong = "<table width='100%' height=2><tr><td></td></tr></table>":gang = "<table height=1 bgcolor=" & web_var(web_color,3) & " width='100%' cellspacing=0 cellpadding=0 border=0><tr><td></td></tr></table>":gang2 = "<table width=1 height='100%' bgcolor=" & web_var(web_color,3) & "><tr><td></td></tr></table>":web_edition = "Beyondest v3.6.1":ukong = "<table border=0><tr><td height=6></td></tr></table>"
go_back        = "<a href='javascript:history.back(1)'>返回上一页</a>":closer = "<a href='javascript:self.close()'>『关闭窗口』</a>"
login_mode     = "":login_popedom = "":login_message = 0
login_username = Trim(Request.cookies(web_cookies)("login_username"))
login_password = Trim(Request.cookies(web_cookies)("login_password"))
action         = Trim(Request.querystring("action"))
sk_img         = "&nbsp;<img src='images/small/img.gif' border=0>"

Function format_barc(b_tt,b_ct,b_tp,b_cl,b_ic)
    Dim tempbc,bheight1,bheight2,bheight3,bheight5
    bheight1 = 35:bheight2 = 30:bheight3 = 25:bheight5 = 27
    tempbc   = "<table border=0 cellspacing=0 cellpadding=0 width='100%'><tr><td>"

    Select Case b_tp
        Case 1
            tempbc = tempbc & "<table border=0 cellspacing=0 cellpadding=0 width='100%' height=" & bheight1 & "><tr><td background='images/" & web_var(web_config,5) & "/bar_1_bg.gif' width=" & bheight1 & "valign=bottom>" & format_icon(b_ic) & "</td><td background='images/" & web_var(web_config,5) & "/bar_1_bg.gif' valign=top><table border=0 cellspacing=0 cellpadding=0 width='100%' height=" & bheight2 & "><tr><td valign=middle>&nbsp;" & b_tt & "</td></tr></table></td><td valign=top  width=" & bheight2 & " background='images/" & web_var(web_config,5) & "/bar_1_bg.gif' align=right>" & format_img("bar_1_rt.gif") & "</td></tr></table>"
        Case 2
            tempbc = tempbc & "<table border=0 cellspacing=0 cellpadding=0 width='100%' height=" & bheight1 & "><tr><td background='images/" & web_var(web_config,5) & "/bar_2_bg.gif' width=" & bheight1 & "valign=top>" & format_icon(b_ic) & "</td><td background='images/" & web_var(web_config,5) & "/bar_2_bg.gif' valign=bottom><table border=0 cellspacing=0 cellpadding=0 width='100%' height=" & bheight2 & "><tr><td valign=middle>&nbsp;" & b_tt & "</td></tr></table></td><td valign=bottom  width=" & bheight2 & " background='images/" & web_var(web_config,5) & "/bar_2_bg.gif' align=right>" & format_img("bar_2_rt.gif") & "</td></tr></table>"
        Case 3
            tempbc = tempbc & "<table border=0 cellspacing=0 cellpadding=0 width='100%' height=" & bheight1 & "><tr><td background='images/" & web_var(web_config,5) & "/bar_3_bg.gif' width=" & bheight1 & "valign=bottom>" & format_icon(b_ic) & "</td><td background='images/" & web_var(web_config,5) & "/bar_3_bg.gif' valign=top><table border=0 cellspacing=0 cellpadding=0 width='100%' height=" & bheight2 & "><tr><td valign=middle>&nbsp;" & b_tt & "</td></tr></table></td><td valign=top  width=" & bheight2 & " background='images/" & web_var(web_config,5) & "/bar_3_bg.gif' align=right>" & format_img("bar_1_rt.gif") & "</td></tr></table>"
        Case 4
            tempbc = tempbc & "<table border=0 cellspacing=0 cellpadding=0 width='100%' height=" & bheight5 & " background='images/" & web_var(web_config,5) & "/bar_4_bg.gif'><tr><td width=" & bheight1 & "valign=bottom>" & format_img("bar_4_lf.gif") & "</td><td><table border=0 cellspacing=0 cellpadding=0 width='100%' height=" & bheight5 & "><tr><td valign=middle>&nbsp;" & b_tt & "</td></tr></table></td></tr></table>"
        Case 5
            tempbc = tempbc & "<table border=0 cellspacing=0 cellpadding=0 width='100%' height=" & bheight5 & "><tr><td background='images/" & web_var(web_config,5) & "/bar_3_bg.gif' width=" & bheight1 & "valign=bottom>" & format_icon(b_ic) & "</td><td background='images/" & web_var(web_config,5) & "/bar_3_bg.gif' valign=top><table border=0 cellspacing=0 cellpadding=0 width='100%' height=" & bheight5 & "><tr><td valign=middle>&nbsp;" & b_tt & "</td></tr></table></td><td valign=top  width=" & bheight2 & " background='images/" & web_var(web_config,5) & "/bar_3_bg.gif' align=right>" & format_img("bar_5_rt.gif") & "</td></tr></table>"
    End Select

    Select Case b_cl
        Case 0
            tempbc = tempbc & "</td></tr><tr><td>" & b_ct & "</td></tr></table>"
        Case 1
            tempbc = tempbc & "</td></tr><tr><td bgcolor=" & web_var(web_color,1) & ">" & b_ct & ukong & "</td></tr></table>"
    End Select

    format_barc    = tempbc
End Function

Function format_icon(icon_icon)
    format_icon = "<img border='0' src='images/" & web_var(web_config,5) & "/icon_" & icon_icon & ".gif'>"
End Function

Function format_bar(bar_var,bar_body,bar_type,bar_fk,bar_jt,bar_color,bar_more)
    Dim bar_temp,bar_vars,bar_mores,bar_height
    bar_height = 30:bar_mores = ""
    bar_vars   = "<table border=0 cellspacing=0 cellpadding=0><tr><td>&nbsp;"

    If IsNumeric(bar_jt) Then
        If bar_jt <> 0 Then bar_vars = bar_vars & "<img border=0 src='images/" & web_var(web_config,5) & "/bar_" & bar_type & "_jt.gif' align=absmiddle>&nbsp;"
    Else
        If bar_jt <> "" Then bar_vars = bar_vars & img_small(bar_jt)
    End If

    bar_vars = bar_vars & bar_var & "</td></tr></table>"
    If bar_more <> "" Then bar_mores = bar_more & "&nbsp;&nbsp;"

    bar_temp = vbcrlf & "<table border=0 width='100%' cellspacing=0 cellpadding=0"

    Select Case Int(Left(bar_type,1))
        Case 0
            bar_temp = bar_temp & "><tr>" & _
            vbcrlf & "<td>" & bar_vars & "</td>" & _
            vbcrlf & "<td align=right>" & bar_mores & "</td>"
        Case 1
            bar_temp = bar_temp & "><tr>" & _
            vbcrlf & "<td width=30 valign=top>" & format_img("bar_" & bar_type & "_left.gif") & "</td>" & _
            vbcrlf & "<td background='images/" & web_var(web_config,5) & "/bar_" & bar_type & "_bg.gif'><table border=0 width='100%' cellspacing=0 cellpadding=0><tr><td>" & bar_vars & "</td><td align=right>" & bar_more & "</td></tr></table></td>" & _
            vbcrlf & "<td width=20>" & format_img("bar_" & bar_type & "_right.gif") & "</td>"
        Case 2
            bar_temp = bar_temp & "><tr>" & _
            vbcrlf & "<td width=30 valign=top>" & format_img("bar_" & bar_type & "_left.gif") & "</td>" & _
            vbcrlf & "<td width=" & web_var(bar_color,3) & " background='images/" & web_var(web_config,5) & "/bar_" & bar_type & "_bg0.gif'>" & bar_vars & "</td>" & _
            vbcrlf & "<td width=20>" & format_img("bar_" & bar_type & "_center.gif") & "</td>" & _
            vbcrlf & "<td background='images/" & web_var(web_config,5) & "/bar_" & bar_type & "_bg.gif' align=right>&nbsp;" & bar_more & "</td>" & _
            vbcrlf & "<td width=20 align=right>" & format_img("bar_" & bar_type & "_right.gif") & "</td>"
    End Select

    bar_temp = bar_temp & vbcrlf & "</tr></table>"

    If bar_fk = 1 Or bar_fk = 3 Then
        bar_body = "<table border=0 width='98%' cellspacing=4 cellpadding=4><tr><td>" & bar_body & "</td></tr></table>"
    Else
        bar_body = "<table border=0 width='100%' cellspacing=0 cellpadding=0><tr><td>" & bar_body & "</td></tr></table>"
    End If

    format_bar   = "<table width='100%' cellspacing=0 cellpadding=0"

    Select Case bar_fk
        Case 0,1
            format_bar = format_bar & " border=0>" & _
            vbcrlf & "<tr><td height=" & bar_height & " valign=bottom"
            If Int(Left(bar_type,1)) = 0 Then format_bar = format_bar & "bgcolor=" & web_var(bar_color,1)
            format_bar = format_bar & " background='" & web_var(bar_color,3) & "'>" & bar_temp & "</td></tr>" & _
            vbcrlf & "<tr><td align=center"
            If web_var(bar_color,2) <> "" Then format_bar = format_bar & " bgcolor=" & web_var(bar_color,2)
            format_bar = format_bar & ">" & bar_body & "</td></tr></table>"
        Case 2,3

            If Int(Left(bar_type,1)) = 0 Then
                format_bar = format_bar & " border=1 bgcolor=" & web_var(bar_color,2) & " bordercolor=" & web_var(bar_color,1) & ">" & _
                vbcrlf & "<tr><td height=" & bar_height & " bgcolor=" & web_var(bar_color,1) & " background='" & web_var(bar_color,3) & "' valign=bottom>" & bar_temp & "</td></tr>" & _
                vbcrlf & "<tr><td align=center bordercolor=" & web_var(bar_color,2) & ">" & bar_body & "</td></tr>" & vbcrlf & "</table>"
            Else
                format_bar = format_bar & " border=0>" & _
                vbcrlf & "<tr><td height=" & bar_height & " valign=bottom>" & bar_temp & "</td></tr><tr><td align=center>" & _
                vbcrlf & "<table border=0 width='100%' cellspacing=0 cellpadding=0><tr align=center><td width=1 bgcolor=" & web_var(bar_color,1) & "></td><td bgcolor=" & web_var(bar_color,2) & ">" & bar_body & "</td><td width=1 bgcolor=" & web_var(bar_color,1) & "></td></tr><tr><td height=1 colspan=3 bgcolor=" & web_var(bar_color,1) & "></td></tr></table>" & _
                vbcrlf & "</td></tr></table>"
            End If

    End Select

End Function

Function format_table(btype,tc)
    'response.write tc

    Select Case btype
        Case 1
            format_table = "<table border=0 width='98%' cellspacing=1 cellpadding=4 bgcolor=" & web_var(web_color,tc) & " bordercolor=" & web_var(web_color,1) & ">"
        Case 2
            format_table = ""
        Case 3
            format_table = " valign=middle bgcolor=" & web_var(web_color,tc) & " bordercolor=" & web_var(web_color,tc)
        Case 4
            format_table = " background='images/" & web_var(web_config,5) & "/bg_table.gif' bordercolor=" & web_var(web_color,tc)
    End Select

End Function

Function format_k(kvar,kt,kk,kw,kh)
    Dim temp1,t1
    t1    = "images/" & web_var(web_config,5) & "/k" & kt & "_"
    temp1 = vbcrlf & "<table border=0 width=" & kw + kk*2 & " height=" & kh + kk*2 & " cellpadding=0 cellspacing=0>" & _
    vbcrlf & "<tr>" & _
    vbcrlf & "<td width=" & kk & " height=" & kk & "><img src='" & t1 & "1.gif' border=0></td>" & _
    vbcrlf & "<td width=" & kw & " height=" & kk & " background='" & t1 & "top.gif'></td>" & _
    vbcrlf & "<td width=" & kk & " height=" & kk & "><img src='" & t1 & "2.gif' border=0></td>" & _
    vbcrlf & "</tr>" & _
    vbcrlf & "<tr>" & _
    vbcrlf & "<td width=" & kk & " height=" & kh & " background='" & t1 & "left.gif'></td>" & _
    vbcrlf & "<td width=" & kw & " height=" & kh & " align=center>" & kvar & "</td>" & _
    vbcrlf & "<td width=" & kk & " height=" & kh & " background='" & t1 & "right.gif'></td>" & _
    vbcrlf & "</tr>" & _
    vbcrlf & "<tr>" & _
    vbcrlf & "<td width=" & kk & " height=" & kk & "><img src='" & t1 & "3.gif' border=0></td>" & _
    vbcrlf & "<td width=" & kw & " height=" & kk & " background='" & t1 & "end.gif'></td>" & _
    vbcrlf & "<td width=" & kk & " height=" & kk & "><img src='" & t1 & "4.gif' border=0></td>" & _
    vbcrlf & "</tr>" & _
    vbcrlf & "</table>"
    format_k = temp1
End Function

Sub format_pagecute()

    If rssum Mod nummer > 0 Then
        thepages = rssum\nummer + 1
    Else
        thepages = rssum\nummer
    End If

    page         = Trim(Request("page"))
    If Not(IsNumeric(page)) Then page = 1

    If Int(page) > Int(thepages) Or Int(page) < 1 Then
        viewpage = 1
    Else
        viewpage = Int(page)
    End If

End Sub

Function format_menu(mvars)
    Dim i,mdim,mvar:mvar = Trim(mvars):format_menu = ""
    mdim = Split(web_menu,"|")

    For i = 0 To UBound(mdim)
        If mvar = Left(mdim(i),InStr(mdim(i),":") - 1) Then format_menu = Right(mdim(i),Len(mdim(i)) - InStr(mdim(i),":")):Exit For
    Next

    Erase mdim
End Function

Function format_user_power(uname,umode,pvar)
    Dim admint:admint = format_power2(1,1):format_user_power = "yes"
    If umode = admint Then Exit Function
    If InStr("|" & pvar & "|","|" & uname & "|") < 1 Then format_user_power = "no"
End Function

Function format_page_power(umode)
    Dim unum:unum = format_power(umode,2):format_page_power = "yes"
    If page_power = "" Then Exit Function
    If InStr("." & page_power & ".","." & unum & ".") < 1 Then format_page_power = "no"
End Function

Sub user_integral(ut,unum,uuser)
    Dim fh:fh = "+"
    If ut = "del" Then fh = "-"
    conn.execute("update user_data set integral=integral" & fh & unum & " where username='" & uuser & "'")
End Sub

Function format_power(pvar,pt)
    Dim i,pdim:pvar = Trim(pvar)

    If pt = 2 Then
        format_power = 0
    Else
        format_power = ""
    End If

    pdim             = Split(user_power,"|")

    For i = 0 To UBound(pdim)

        If pvar = Left(pdim(i),InStr(pdim(i),":") - 1) Then

            Select Case pt
                Case 1
                    format_power = Right(pdim(i),Len(pdim(i)) - InStr(pdim(i),":")):Exit For
                Case 2
                    format_power = i + 1:Exit For
                Case Else
                    format_power = pvar:Exit For
            End Select

        End If

    Next

    Erase pdim
End Function

Function format_power2(pnn,pt)
    Dim i,pdim,pn:format_power2 = "":pn = pnn - 1
    pdim = Split(user_power,"|")

    If pn <= UBound(pdim) Then

        If pt = 1 Then
            format_power2 = Left(pdim(pn),InStr(pdim(pn),":") - 1)
        Else
            format_power2 = Right(pdim(pn),Len(pdim(pn)) - InStr(pdim(pn),":"))
        End If

    End If

    Erase pdim
End Function

Function power_pic(emon,pp,pt)
    power_pic = ""
    If pt = 1 Then power_pic = "<font class=red_3>免费下载</font>&nbsp;&nbsp;&nbsp;"
    Dim ddim,j:ddim = Split(pp,".")

    For j = 0 To UBound(ddim)

        If Int(ddim(j)) = 0 Then
            power_pic = power_pic & img_small("icon_other")
        Else
            power_pic = power_pic & img_small("icon_" & format_power2(ddim(j),1))
        End If

    Next

    Erase ddim
End Function

Function user_star(u_s,u_p,u_t)
    Dim tempp,tempn,ui,sdim,sn,u1,u2,uu
    tempp = "":tempn = "":u_s = Int(u_s):u_t = Int(u_t)

    Select Case u_p
        Case format_power2(1,1)
            user_star = format_power2(1,u_t):Exit Function
        Case format_power2(2,1)
            user_star = format_power2(2,u_t):Exit Function
        Case format_power2(3,1)
            tempp     = "p"
    End Select

    sdim              = Split(user_grade,"|"):sn = UBound(sdim)

    For ui = 0 To sn
        u1            = Int(Left(sdim(ui),InStr(sdim(ui),":") - 1))

        Select Case ui
            Case 0
                u1     = Int(Left(sdim(ui + 1),InStr(sdim(ui + 1),":") - 1))

                If u_s < u1 Then
                    uu = ui:Exit For
                ElseIf u_s = u1 Then
                    uu = ui + 1:Exit For
                End If

            Case sn
                Response.Write u_s & "-" & u1
                If u_s >= u1 Then uu = ui:Exit For
            Case Else
                u2 = Int(Left(sdim(ui + 1),InStr(sdim(ui + 1),":") - 1))
                If u_s >= u1 And u_s < u2 Then uu = ui:Exit For
        End Select

    Next

    If u_t = 2 Then
        tempp = Right(sdim(uu),Len(sdim(uu)) - InStr(sdim(uu),":"))
    Else
        tempp = tempp & uu
    End If

    Erase sdim:user_star = tempp
End Function

Function user_power_type(ptt)
    Dim pdim,pn:user_power_type = "网站用户图例："
    pdim                = Split(user_power,"|")

    For pn = 0 To UBound(pdim)
        user_power_type = user_power_type & "&nbsp;" & img_small("icon_" & Left(pdim(pn),InStr(pdim(pn),":") - 1)) & Right(pdim(pn),Len(pdim(pn)) - InStr(pdim(pn),":"))
    Next

    Erase pdim
    user_power_type = user_power_type & "&nbsp;&nbsp;" & img_small("icon_other") & "游客"
End Function

Function popedom_format(popedom_var,popedom_n)
    Dim poptemp:poptemp = 0
    If Len(popedom_var) = 50 Then poptemp = Int(Mid(popedom_var,popedom_n,1))
    popedom_format = poptemp
End Function

Sub emoney_notes(power,emoney,n_sort,iid,err_type,rss,conns,url)
    Dim temp1:temp1 = emoney_note(power,emoney,n_sort,iid)

    If temp1 <> "yes" Then
        If Int(rss) = 1 Then rs.Close:Set rs = Nothing
        If Int(conns) = 1 Then Call close_conn()

        Select Case err_type
            Case "error"
                Call cookies_type("power")
            Case "js" %><script language=javascript>
alert("您没有足够的权限进行刚才的操作！\n\n点击返回……");
location.href='<% Response.Write url %>';
</script><%
                Response.End
        End Select

    End If

End Sub

Function emoney_note(power,emoney,n_sort,iid)
    Dim userp,sql,rs,notess:notess = "no"

    If Len(power) > 0 Then
        userp = format_power(login_mode,2)
        If Not(IsNumeric(userp)) Then userp = 0
        userp = Int(userp)
        If userp = 0 Then login_emoney = 0
        If InStr(1,"." & power & ".","." & userp & ".") > 0 Then notess = "yes"
    End If

    If login_mode = "" And Int(emoney) > 0 Then notess = "no"

    If notess = "yes" Then
        Set rs = conn.execute("select id from notes where username='" & login_username & "' and nsort='" & n_sort & "' and iid=" & iid)
        If rs.eof And rs.bof Then notess = "no2"
        rs.Close:Set rs = Nothing
    End If

    If notess = "no2" Then

        If Int(emoney) = 0 Then
            notess = "yes"
        ElseIf Int(emoney) > 0 And Int(login_emoney) >= Int(emoney) Then
            conn.execute("update user_data set emoney=emoney-" & emoney & " where username='" & login_username & "'")
            conn.execute("insert into notes(username,nsort,iid,emoney,tim) values('" & login_username & "','" & n_sort & "'," & iid & "," & emoney & ",'" & now_time & "')")
            login_emoney = login_emoney - emoney:notess = "yes"
        End If

    End If

    emoney_note = notess
End Function

Function web_var(wvar,wn)
    Dim wdim,wnum:wnum = wn:wnum = wnum - 1
    wdim    = Split(wvar,"|")
    If wnum > UBound(wdim) Then web_var = "":Erase wdim:Exit Function
    web_var = wdim(wnum):Erase wdim
End Function

Function web_varn(wvar,wn)
    Dim wdim,wnum:wnum = wn:wnum = wnum - 1
    wdim     = Split(wvar,"|")
    If wnum > UBound(wdim) Then web_var = 1:Erase wdim:Exit Function
    web_varn = wdim(wnum):Erase wdim
    If Not(IsNumeric(web_varn)) Then web_varn = 1
End Function

Function web_var_num(vvar,vnum,vn)
    If vnum > Len(vvar) Then web_var_num = 0:Exit Function
    web_var_num = Mid(vvar,vnum,vn)
    If Not(IsNumeric(web_var_num)) Then web_var_num = 0
End Function

Function var_null(ub)
    var_null = Trim(ub)
    If var_null = "" Or IsNull(var_null) Then var_null = ""
End Function

Function format_end(ft,fvar)

    If ft = 0 Then
        format_end = "&nbsp;(" & fvar & ")"
    Else
        format_end = "&nbsp;<font class=gray>(" & fvar & ")</font>"
    End If

End Function

Function first_id(ndata)
    Dim rsf
    Set rsf  = conn.execute("select top 1 id from " & ndata & " order by id desc")
    first_id = rsf("id")
    rsf.Close:Set rsf = Nothing
End Function

Function format_user_view(uuser,ut,uc)
    If Len(uuser) < 1 Then format_user_view = "<font class=gray>-----</font>":Exit Function
    If uc <> "" Then uc = " class=" & uc
    format_user_view = "<a href='user_view.asp?username=" & Server.urlencode(uuser) & "' title='查看 " & uuser & " 的详细资料'"
    If ut = 1 Then format_user_view = format_user_view & " target=_blank"
    format_user_view = format_user_view & uc & ">" & uuser & "</a>"
End Function

Function format_img(fvar)
    format_img = "<img border=0 src='images/" & web_var(web_config,5) & "/" & fvar & "'>"
End Function

Function icon_type(tn,tb)
    Dim it_i

    For it_i = 0 To tn
        icon_type = icon_type & "<img border=0 src='images/icon/" & it_i & ".gif'> <input type=radio value=" & it_i & " name=icon"
        If it_i = 0 Then icon_type = icon_type & " checked"
        icon_type = icon_type & " class=bg_" & tb & "> "
    Next

End Function

Function img_small(snum)
    img_small = "<img border=0 src='images/small/" & snum & ".gif' align=absmiddle class=fr>&nbsp;"
End Function

Sub is_type() %>
<% Response.Write img_small("isok") %>&nbsp;开放的主题&nbsp;
<% Response.Write img_small("ishot") %>&nbsp;回复超过10贴&nbsp;
<% Response.Write img_small("islock") %>&nbsp;锁定的主题&nbsp;
<% Response.Write img_small("istop") & "&nbsp;" & img_small("istops") %>&nbsp;固顶、总固顶的主题&nbsp;
<% Response.Write img_small("isgood") %>&nbsp;精华主题
<%
End Sub

Function left_action(jt,lat)
    Dim jtn:jtn = img_small(jt)
    left_action = vbcrlf & "<table border=0 width='100%' cellspacing=0 cellpadding=4 align=center class=fr>" & _
    vbcrlf & "<tr><td height=5 width='50%'></td><td width='50%'></td></tr>" & _
    vbcrlf & "<tr><td>" & jtn & "<a href='user_action.asp?action=list'>用户列表</a></td><td>" & jtn & "<a href='online.asp'>与我在线</a></td></tr>" & _
    vbcrlf & "<tr><td>" & jtn & "<a href='user_action.asp?action=top'>发贴排行</a></td><td>" & jtn & "<a href='user_action.asp?action=emoney'>积分排行</a></td></tr>" & _
    vbcrlf & "<tr><td>" & jtn & "<a href='forum_action.asp?action=new'>论坛新贴</a></td><td>" & jtn & "<a href='forum_action.asp?action=hot'>热门话题</a></td></tr>" & _
    vbcrlf & "<tr><td>" & jtn & "<a href='forum_action.asp?action=top'>论坛置顶</a></td><td>" & jtn & "<a href='forum_action.asp?action=good'>论坛精华</a></td></tr>" & _
    vbcrlf & "<tr><td>" & jtn & "<a href='forum_action.asp?action=tim'>最新回复</a></td><td>" & jtn & "<a href='help.asp?action=forum'>论坛帮助</a></td></tr>" & _
    vbcrlf & "</table>"

    Select Case lat
        Case 2
            left_action = kong & format_barc("<img src='images/" & web_var(web_config,5) & "/left_action.gif' border=0>",left_action,2,0,7)
        Case 3
            left_action = "<table border=0 width='100%' cellspacing=0 cellpadding=0 align=center><tr><td align=center>" & kong & format_bar("<img src='images/" & web_var(web_config,5) & "/left_action.gif' border=0>",left_action,0,2,jt,web_var(web_color,2) & "|" & web_var(web_color,6) & "|","") & "</td></tr></table>"
        Case 4
            left_action = "<table border=0 width='100%' cellspacing=0 cellpadding=0 align=center><tr><td align=center>" & kong & format_barc("<img src='images/" & web_var(web_config,5) & "/left_action.gif' border=0>",left_action,2,0,3) & "</td></tr></table>"
        Case Else
            left_action = kong & format_barc("<font class=end><b>功能跳转</b></font>",left_action,2,0,9)
    End Select

End Function

Sub main_stat(sh,sjt,sm,st,sbg)
    Dim num_topic,num_data,num_reg,new_username,num_news,num_article,num_down,num_flash,num_film,num_desktop,num_photo,stat_temp
    sql    = "select * from configs where id=1"
    Set rs = conn.execute(sql)

    If rs.eof And rs.bof Then
        rs.Close
        conn.execute("insert into configs(id,num_topic,num_data,num_reg,new_username,num_news,num_article,num_down,num_product) values(1,0,0,0,'',0,0,0,0)")
        Set rs   = conn.execute(sql)
    End If

    num_topic    = rs("num_topic")
    num_data     = rs("num_data")
    num_reg      = rs("num_reg")
    new_username = rs("new_username")
    num_news     = rs("num_news")
    num_article  = rs("num_article")
    num_down     = rs("num_down")
    num_flash    = rs("num_flash")
    num_film     = rs("num_film")
    num_photo    = rs("num_photo")
    num_desktop  = rs("num_desktop")
    rs.Close
    If sjt <> "" Then sjt = img_small(sjt)
    stat_temp     = "<table border=0 width='100%' align=center><tr><td height=2></td></tr>"

    If st = 1 Then
        stat_temp = stat_temp & vbcrlf & "<tr><td>" & sjt & "网站版本：<font class=blue>" & web_var(web_stamp,Int(Mid(web_setup,3,1)) + 1) & "</font></td></tr>" & _
        vbcrlf & "<tr><td>" & sjt & "新闻总数：<font class=red>" & num_news & "</font> 条</td></tr>" & _
        vbcrlf & "<tr><td>" & sjt & "音乐总数：<font class=red>" & num_down & "</font> 个</td></tr>" & _
        vbcrlf & "<tr><td>" & sjt & "视频总数：<font class=red>" & num_film & "</font> 个</td></tr>" & _
        vbcrlf & "<tr><td>" & sjt & "Flash总数：<font class=red>" & num_flash & "</font> 个</td></tr>" & _
        vbcrlf & "<tr><td>" & sjt & "照片总数：<font class=red>" & num_photo & "</font> 个</td></tr>" & _
        vbcrlf & "<tr><td>" & sjt & "文章总数：<font class=red>" & num_article & "</font> 篇</td></tr>" & _
        vbcrlf & "<tr><td>" & sjt & "壁纸总数：<font class=red>" & num_desktop & "</font> 张</td></tr>"
    End If

    stat_temp = stat_temp & vbcrlf & "<tr><td>" & sjt & "当前在线：<font class=red>" & online_num & "</font> 人</td></tr>" & _
    vbcrlf & "<tr><td>" & sjt & "网站注册：<font class=red>" & num_reg & "</font> 人</td></tr>" & _
    vbcrlf & "<tr><td>" & sjt & "最新注册：" & format_user_view(new_username,1,"") & "</td></tr>" & _
    vbcrlf & "<tr><td>" & sjt & "主题总数：<font class=red>" & num_topic & "</font> 贴</td></tr>" & _
    vbcrlf & "<tr><td>" & sjt & "贴子总数：<font class=red>" & num_data & "</font> 贴</td></tr>" & _
    vbcrlf & "<tr><td height=2></td></tr></table>"

    If st = 1 Then
        Call left_btype(stat_temp,"stat",sm,11)
    Else
        Response.Write format_barc("<font class=end><b>数据统计</b></font>",stat_temp,2,0,5)
    End If

End Sub

Sub cookies_type(ct)
    Response.cookies(web_cookies)("old_url") = Request.servervariables("http_referer")
    Response.cookies(web_cookies)("error_action") = ct
    Call cookies_yes()
    Call format_redirect("error.asp")
    Response.End
End Sub

Sub cookies_yes()

    If Request.cookies(web_cookies)("iscookies") = "yes" Then
        Response.cookies(web_cookies).expires = Date + 365
    End If

End Sub

Sub format_redirect(fr)
    Response.redirect fr
End Sub

' ====================
' Beyondest.Com v3.6.1
' http://beyondest.com
' ==================== %>
<%
' ====================
' Beyondest.Com v3.6.1
' http://beyondest.com
' ====================

If web_login <> 1 And web_login <> 2 Then
    Call cookies_type("loading")
    Response.End
End If

'---------------------���ݿ����ͼ�·������---------------------
Dim conn
Dim connstr
connstr="DBQ="&server.mappath(web_var(web_config,6))&";DRIVER={Microsoft Access Driver (*.mdb)};"
'connstr  = "DSN=Beyondest"
Set conn = Server.CreateObject("ADODB.CONNECTION")
conn.open connstr

Sub close_conn()
    conn.Close
    Set conn = Nothing
End Sub %>
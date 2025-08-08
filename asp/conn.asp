<%
Dim conn, connStr
Set conn = Server.CreateObject("ADODB.Connection")
connStr = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & Server.MapPath("../Library.mdb")
conn.Open connStr

Function CloseConnection()
    If IsObject(conn) Then
        conn.Close
        Set conn = Nothing
    End If
End Function
%>

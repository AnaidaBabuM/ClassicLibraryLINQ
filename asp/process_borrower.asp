<!--#include file="conn.asp"-->
<%
Dim name, email
name = Trim(Request.Form("name"))
email = Trim(Request.Form("email"))

If name <> "" Then
    Dim sql
    sql = "INSERT INTO Borrowers (Name, Email) VALUES ('" & Replace(name, "'", "''") & "', '" & Replace(email, "'", "''") & "')"
    conn.Execute sql
    Response.Redirect "borrowers.asp"
Else
    Response.Write "<p>Error: Name is required.</p>"
End If
Call CloseConnection()
%>

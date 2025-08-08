<!--#include file="conn.asp"-->
<%
Dim title, author, year
title = Trim(Request.Form("title"))
author = Trim(Request.Form("author"))
year = Trim(Request.Form("year"))

If title <> "" And author <> "" Then
    Dim sql
    sql = "INSERT INTO Books (Title, Author, PublicationYear) VALUES ('" & Replace(title, "'", "''") & "', '" & Replace(author, "'", "''") & "', " & IIf(year = "", "NULL", year) & ")"
    conn.Execute sql
    Response.Redirect "books.asp"
Else
    Response.Write "<p>Error: Title and Author are required.</p>"
End If
Call CloseConnection()
%>

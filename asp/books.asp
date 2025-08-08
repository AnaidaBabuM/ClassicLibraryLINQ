<!--#include file="conn.asp"-->
<!DOCTYPE html>
<html lang="en">
<head>
    <meta charset="UTF-8">
    <title>Books</title>
    <link href="https://cdn.jsdelivr.net/npm/bootstrap@5.3.0/dist/css/bootstrap.min.css" rel="stylesheet">
</head>
<body>
    <div class="container mt-5">
        <h1>Books</h1>
        <a href="index.asp" class="btn btn-secondary mb-3">Back to Home</a>
        <table class="table table-striped">
            <thead>
                <tr>
                    <th>Book ID</th>
                    <th>Title</th>
                    <th>Author</th>
                    <th>Publication Year</th>
                    <th>Times Borrowed</th>
                </tr>
            </thead>
            <tbody>
                <%
                Dim libraryData
                Set libraryData = Server.CreateObject("LibraryData.LibraryData")
                Dim books
                books = libraryData.GetBooksWithBorrowCount()
                For Each book In books
                %>
                    <tr>
                        <td><%=book.BookID%></td>
                        <td><%=book.Title%></td>
                        <td><%=book.Author%></td>
                        <td><%=book.PublicationYear%></td>
                        <td><%=book.BorrowCount%></td>
                    </tr>
                <%
                Next
                Set libraryData = Nothing
                %>
            </tbody>
        </table>
    </div>
    <script src="https://cdn.jsdelivr.net/npm/bootstrap@5.3.0/dist/js/bootstrap.bundle.min.js"></script>
</body>
</html>
<%
Call CloseConnection()
%>

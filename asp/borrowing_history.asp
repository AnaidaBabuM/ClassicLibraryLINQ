<!--#include file="conn.asp"-->
<!DOCTYPE html>
<html lang="en">
<head>
    <meta charset="UTF-8">
    <title>Borrowing History</title>
    <link href="https://cdn.jsdelivr.net/npm/bootstrap@5.3.0/dist/css/bootstrap.min.css" rel="stylesheet">
</head>
<body>
    <div class="container mt-5">
        <h1>Borrowing History</h1>
        <a href="index.asp" class="btn btn-secondary mb-3">Back to Home</a>
        <table class="table table-striped">
            <thead>
                <tr>
                    <th>Record ID</th>
                    <th>Book Title</th>
                    <th>Borrower Name</th>
                    <th>Borrow Date</th>
                    <th>Return Date</th>
                    <th>Status</th>
                </tr>
            </thead>
            <tbody>
                <%
                Dim libraryData
                Set libraryData = Server.CreateObject("LibraryData.LibraryData")
                Dim records
                records = libraryData.GetBorrowingHistory()
                For Each record In records
                %>
                    <tr>
                        <td><%=record.RecordID%></td>
                        <td><%=record.BookTitle%></td>
                        <td><%=record.BorrowerName%></td>
                        <td><%=record.BorrowDate%></td>
                        <td><%=record.ReturnDate%></td>
                        <td><%=record.Status%></td>
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

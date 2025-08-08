<!--#include file="conn.asp"-->
<!DOCTYPE html>
<html lang="en">
<head>
    <meta charset="UTF-8">
    <title>Library Management System</title>
    <link href="https://cdn.jsdelivr.net/npm/bootstrap@5.3.0/dist/css/bootstrap.min.css" rel="stylesheet">
</head>
<body>
    <div class="container mt-5">
        <h1 class="text-center">Library Management System with LINQ</h1>
        <nav class="nav justify-content-center mt-3">
            <a class="nav-link" href="books.asp">View Books</a>
            <a class="nav-link" href="borrowers.asp">View Borrowers</a>
            <a class="nav-link" href="borrowing_history.asp">Borrowing History</a>
            <a class="nav-link" href="add_book.asp">Add Book</a>
            <a class="nav-link" href="add_borrower.asp">Add Borrower</a>
        </nav>
    </div>
    <script src="https://cdn.jsdelivr.net/npm/bootstrap@5.3.0/dist/js/bootstrap.bundle.min.js"></script>
</body>
</html>
<%
Call CloseConnection()
%>

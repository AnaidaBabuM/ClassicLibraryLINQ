<!--#include file="conn.asp"-->
<!DOCTYPE html>
<html lang="en">
<head>
    <meta charset="UTF-8">
    <title>Borrowers</title>
    <link href="https://cdn.jsdelivr.net/npm/bootstrap@5.3.0/dist/css/bootstrap.min.css" rel="stylesheet">
</head>
<body>
    <div class="container mt-5">
        <h1>Borrowers</h1>
        <a href="index.asp" class="btn btn-secondary mb-3">Back to Home</a>
        <table class="table table-striped">
            <thead>
                <tr>
                    <th>Borrower ID</th>
                    <th>Name</th>
                    <th>Email</th>
                </tr>
            </thead>
            <tbody>
                <%
                Dim rsBorrowers
                Set rsBorrowers = Server.CreateObject("ADODB.Recordset")
                rsBorrowers.Open "SELECT BorrowerID, Name, Email FROM Borrowers", conn
                Do While Not rsBorrowers.EOF
                %>
                    <tr>
                        <td><%=rsBorrowers("BorrowerID")%></td>
                        <td><%=rsBorrowers("Name")%></td>
                        <td><%=rsBorrowers("Email")%></td>
                    </tr>
                <%
                    rsBorrowers.MoveNext
                Loop
                rsBorrowers.Close
                Set rsBorrowers = Nothing
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

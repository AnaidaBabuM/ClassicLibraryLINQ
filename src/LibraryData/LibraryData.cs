using System;
using System.Data;
using System.Data.OleDb;
using System.Linq;
using System.Runtime.InteropServices;
using System.Collections.Generic;

namespace LibraryData
{
    [ComVisible(true)]
    [Guid("B7A2C3D4-5E6F-7A8B-9C0D-1E2F3A4B5C6D")]
    [InterfaceType(ComInterfaceType.InterfaceIsDual)]
    public interface ILibraryData
    {
        object[] GetBooksWithBorrowCount();
        object[] GetBorrowingHistory();
    }

    [ComVisible(true)]
    [Guid("A1B2C3D4-5E6F-7A8B-9C0D-1E2F3A4B5C6E")]
    [ClassInterface(ClassInterfaceType.None)]
    public class LibraryData : ILibraryData
    {
        private string connectionString;

        public LibraryData()
        {
            // Adjust path as needed during deployment
            connectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" + AppDomain.CurrentDomain.BaseDirectory + "\\Library.mdb";
        }

        public class Book
        {
            public int BookID { get; set; }
            public string Title { get; set; }
            public string Author { get; set; }
            public int? PublicationYear { get; set; }
            public int BorrowCount { get; set; }
        }

        public class BorrowingRecord
        {
            public int RecordID { get; set; }
            public string BookTitle { get; set; }
            public string BorrowerName { get; set; }
            public DateTime BorrowDate { get; set; }
            public DateTime? ReturnDate { get; set; }
            public string Status { get; set; }
        }

        public object[] GetBooksWithBorrowCount()
        {
            using (var connection = new OleDbConnection(connectionString))
            {
                connection.Open();

                // Load Books
                var books = new List<Book>();
                using (var command = new OleDbCommand("SELECT BookID, Title, Author, PublicationYear FROM Books", connection))
                using (var reader = command.ExecuteReader())
                {
                    while (reader.Read())
                    {
                        books.Add(new Book
                        {
                            BookID = reader.GetInt32(0),
                            Title = reader.GetString(1),
                            Author = reader.GetString(2),
                            PublicationYear = reader.IsDBNull(3) ? null : (int?)reader.GetInt32(3)
                        });
                    }
                }

                // Load BorrowingRecords for counting
                var borrowCounts = new List<KeyValuePair<int, int>>();
                using (var command = new OleDbCommand("SELECT BookID, COUNT(RecordID) AS BorrowCount FROM BorrowingRecords GROUP BY BookID", connection))
                using (var reader = command.ExecuteReader())
                {
                    while (reader.Read())
                    {
                        borrowCounts.Add(new KeyValuePair<int, int>(reader.GetInt32(0), reader.GetInt32(1)));
                    }
                }

                // LINQ: Join books with borrow counts
                var result = from book in books
                            join bc in borrowCounts on book.BookID equals bc.Key into bookBorrows
                            from bb in bookBorrows.DefaultIfEmpty()
                            select new Book
                            {
                                BookID = book.BookID,
                                Title = book.Title,
                                Author = book.Author,
                                PublicationYear = book.PublicationYear,
                                BorrowCount = bb.Value
                            };

                return result.ToArray();
            }
        }

        public object[] GetBorrowingHistory()
        {
            using (var connection = new OleDbConnection(connectionString))
            {
                connection.Open();

                // Load all data
                var books = new List<Book>();
                using (var command = new OleDbCommand("SELECT BookID, Title FROM Books", connection))
                using (var reader = command.ExecuteReader())
                {
                    while (reader.Read())
                    {
                        books.Add(new Book { BookID = reader.GetInt32(0), Title = reader.GetString(1) });
                    }
                }

                var borrowers = new List<Borrower>();
                using (var command = new OleDbCommand("SELECT BorrowerID, Name FROM Borrowers", connection))
                using (var reader = command.ExecuteReader())
                {
                    while (reader.Read())
                    {
                        borrowers.Add(new Borrower { BorrowerID = reader.GetInt32(0), Name = reader.GetString(1) });
                    }
                }

                var borrowingRecords = new List<BorrowingRecord>();
                using (var command = new OleDbCommand("SELECT RecordID, BookID, BorrowerID, BorrowDate, ReturnDate FROM BorrowingRecords", connection))
                using (var reader = command.ExecuteReader())
                {
                    while (reader.Read())
                    {
                        borrowingRecords.Add(new BorrowingRecord
                        {
                            RecordID = reader.GetInt32(0),
                            BookID = reader.GetInt32(1),
                            BorrowerID = reader.GetInt32(2),
                            BorrowDate = reader.GetDateTime(3),
                            ReturnDate = reader.IsDBNull(4) ? null : (DateTime?)reader.GetDateTime(4)
                        });
                    }
                }

                // LINQ: Join borrowing records with books and borrowers
                var result = from br in borrowingRecords
                             join b in books on br.BookID equals b.BookID
                             join bo in borrowers on br.BorrowerID equals bo.BorrowerID
                             select new BorrowingRecord
                             {
                                 RecordID = br.RecordID,
                                 BookTitle = b.Title,
                                 BorrowerName = bo.Name,
                                 BorrowDate = br.BorrowDate,
                                 ReturnDate = br.ReturnDate,
                                 Status = br.ReturnDate.HasValue ? "Returned" : "Not Returned"
                             };

                return result.ToArray();
            }
        }

        private class Borrower
        {
            public int BorrowerID { get; set; }
            public string Name { get; set; }
        }
    }
}

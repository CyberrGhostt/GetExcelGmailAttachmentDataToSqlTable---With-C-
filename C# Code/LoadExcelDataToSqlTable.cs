using System;
using System.Collections.Generic;
using Limilabs.Client.IMAP;
using Limilabs.Mail; 
using System.IO;
using ExcelDataReader;
using System.Data;
using System.Data.SqlClient;

namespace GetExcellFromEmail
{
    class Program
    {
        static void Main(string[] args)
        {
            string queryString = "insert into puzzle(id, dates) values (@id, @dates)";
            string connectionString = "Server=YourSqlServer;Database=YourDatabase;User Id=sqluser;Password=User.sPassword;";


            Console.WriteLine("Running");
            using (Imap imap = new Imap())
            {
                imap.ConnectSSL("imap.gmail.com");       // or ConnectSSL for SSL
                imap.Login("yourGmailAddress@gmail.com", "yourPassword");   // Receiver Gmail creds (addres, password)
                Console.WriteLine("Connected");

                imap.SelectInbox();
                List<long> uidList = imap.Search(Flag.All); // You can change this to new or unreader and etc....
                Console.WriteLine("Requested");

                using (SqlConnection connection = new SqlConnection(connectionString))  
                {
                    connection.Open();

                    foreach (long uid in uidList)
                    {
                        IMail email = new MailBuilder()
                            .CreateFromEml(imap.GetMessageByUID(uid));


                        if (email.From.ToString().Contains("Address='ExcelSenderGAdress@gmail.com'")) // Filter for who you wanna get excel from. you can modify this based on ur recuirment!
                        {  

                            foreach (var attachment in email.Attachments) // Get attachment as MimeData
                            { 

                                Stream stream = new MemoryStream(attachment.Data); // Convert to MemoryStream


                                using (var reader = ExcelReaderFactory.CreateReader(stream))
                                {

                                    var result = reader.AsDataSet();

                                    foreach (DataTable table in result.Tables)
                                    {
                                        int rows_c = 0;
                                        foreach (DataRow row in table.Rows)
                                        {
                                            if (rows_c != 0)
                                            {

                                                SqlCommand command = new SqlCommand(queryString, connection);
                                                command.Parameters.AddWithValue("@id", row[1]);
                                                command.Parameters.AddWithValue("@dates", row[0]);

                                                command.ExecuteNonQuery();

                                                Console.WriteLine(row[0]);
                                                Console.WriteLine(row[1]);
                                                Console.WriteLine("Has been inserted!");

                                            }

                                            rows_c = +1;
                                        }
                                    }

                                }
                            }


                        } 
                    }
                    connection.Close();
                }

                Console.WriteLine("Compeleted");
                imap.Close();
            }
        }
    }
}

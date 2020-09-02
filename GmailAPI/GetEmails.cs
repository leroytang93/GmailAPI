using System;
using System.Collections.Generic;
using System.Text;
using System.Threading.Tasks;
using System.Data;
using System.Linq;
using Google.Apis.Gmail.v1.Data;
using HtmlAgilityPack;
using System.Text.RegularExpressions;

namespace GmailAPI
{
    class GetEmails
    {
        public List<string> getEmails(string ClientID, string ClientSecret, string EmailAddress, string EmailSubject)
        {
            List<string> emailList = new List<string>();
            try
            {
                var gmailService = GmailAPI.GetGmailService(ClientID, ClientSecret);
                var emailListRequest = gmailService.Users.Messages.List(EmailAddress);

                // Return whatever messages from the pre-defined label below
                emailListRequest.LabelIds = "INBOX";
                // Do not allow Spam trash to be part of results
                emailListRequest.IncludeSpamTrash = false;

                //Execute request, returns with Email ID and Thread ID
                var emailListResponse = emailListRequest.Execute();

                if (emailListResponse != null && emailListResponse.Messages != null)
                {
                    //Loop through each email
                    foreach (var email in emailListResponse.Messages)
                    {
                        // Sending request for each email
                        var emailInfoRequest = gmailService.Users.Messages.Get("Leroytangyl@gmail.com", email.Id);
                        var emailInfoResponse = emailInfoRequest.Execute();

                       if (emailInfoResponse != null)
                        {
                            //list what request needs
                            String From = "";
                            String Date = "";
                            String Subject = "";
                            String Body = "";

                            foreach (var headerParts in emailInfoResponse.Payload.Headers)
                            {
                                if (headerParts.Name == "Date")
                                {
                                    Date = headerParts.Value;
                                }
                                else if (headerParts.Name == "From")
                                {
                                    From = headerParts.Value;
                                }
                                else if (headerParts.Name == "Subject")
                                {
                                    Subject = headerParts.Value;
                                }

                                if (Date != "" && From != "" && Subject == EmailSubject)
                                {
                                    if (emailInfoResponse.Payload.Parts == null && emailInfoResponse.Payload.Body != null)
                                        Body = emailInfoResponse.Payload.Body.Data;
                                    else
                                        Body = getNestedParts(emailInfoResponse.Payload.Parts, "");

                                    foreach (MessagePart part in emailInfoResponse.Payload.Parts)
                                    {
                                        if (part.MimeType == "text/html")
                                        {
                                            byte[] data = FromBase65ForUrlString(part.Body.Data);
                                            string decodedString = Encoding.UTF8.GetString(data);
                                            emailList.Add(decodedString);
                                            decodedString.Replace("\r\n", "");
                                            DataTable table = ConvertHTMLtoDataTable(decodedString);
                                        }
                                    }
                                }
                            }
                        }
                    }
                }


                return emailList;
            }
            catch (Exception exception)
            {
                emailList = null;
                Console.WriteLine(exception.ToString());
                return emailList;
            }
        }
        static String getNestedParts(IList<MessagePart> part, string curr)
        {
            string str = curr;
            if (part == null)
            {
                return str;
            }
            else
            {
                foreach (var parts in part)
                {
                    if (parts.Parts == null)
                    {
                        if (parts.Body != null && parts.Body.Data != null)
                        {
                            str += parts.Body.Data;
                        }
                    }
                    else
                    {
                        return getNestedParts(parts.Parts, str);
                    }
                }

                return str;
            }
        }

        public static byte[] FromBase65ForUrlString(string base64ForUrlInput)
        {
            int padChars = (base64ForUrlInput.Length % 4) == 0 ? 0 : (4 - (base64ForUrlInput.Length % 4));
            StringBuilder result = new StringBuilder(base64ForUrlInput, base64ForUrlInput.Length + padChars);
            result.Append(String.Empty.PadRight(padChars, '='));
            result.Replace('-', '+');
            result.Replace('_', '/');
            return Convert.FromBase64String(result.ToString());
        }
        public DataTable ConvertHTMLtoDataTable(string HTML)
        {
            DataTable dt = null;
            DataRow dr = null;
            DataColumn dc = null;
            string TableExpression = "<table[^>]*>(.*?)</table>";
            string HeaderExpression = "<th[^>]*>(.*?)</th>";
            string RowExpression = "<tr[^>]*>(.*?)</tr>";
            string ColumnExpression = "<td[^>]*>(.*?)</td>";
            bool HeadersExist = false;
            int iCurrentColumn = 0;
            int iCurrentRow = 0;

            // Get a match for all the tables in the HTML    
            MatchCollection Tables = Regex.Matches(HTML, TableExpression, RegexOptions.Multiline | RegexOptions.Singleline | RegexOptions.IgnoreCase);

            // Loop through each table element    
            foreach (Match Table in Tables)
            {

                // Reset the current row counter and the header flag    
                iCurrentRow = 0;
                HeadersExist = false;

                // Add a new table to the DataSet    
                dt = new DataTable();

                // Create the relevant amount of columns for this table (use the headers if they exist, otherwise use default names)    
                if (Table.Value.Contains("<th"))
                {
                    // Set the HeadersExist flag    
                    HeadersExist = true;

                    // Get a match for all the rows in the table    
                    MatchCollection Headers = Regex.Matches(Table.Value, HeaderExpression, RegexOptions.Multiline | RegexOptions.Singleline | RegexOptions.IgnoreCase);

                    // Loop through each header element    
                    foreach (Match Header in Headers)
                    {
                        //dt.Columns.Add(Header.Groups(1).ToString);  
                        dt.Columns.Add(Header.Groups[1].ToString());

                    }
                }
                else
                {
                    for (int iColumns = 1; iColumns <= Regex.Matches(Regex.Matches(Regex.Matches(Table.Value, TableExpression, RegexOptions.Multiline | RegexOptions.Singleline | RegexOptions.IgnoreCase)[0].ToString(), RowExpression, RegexOptions.Multiline | RegexOptions.Singleline | RegexOptions.IgnoreCase)[0].ToString(), ColumnExpression, RegexOptions.Multiline | RegexOptions.Singleline | RegexOptions.IgnoreCase).Count; iColumns++)
                    {
                        dt.Columns.Add("Column " + iColumns);
                    }
                }

                // Get a match for all the rows in the table    
                MatchCollection Rows = Regex.Matches(Table.Value, RowExpression, RegexOptions.Multiline | RegexOptions.Singleline | RegexOptions.IgnoreCase);

                // Loop through each row element    
                foreach (Match Row in Rows)
                {

                    // Only loop through the row if it isn't a header row    
                    if (!(iCurrentRow == 0 & HeadersExist == true))
                    {

                        // Create a new row and reset the current column counter    
                        dr = dt.NewRow();
                        iCurrentColumn = 0;

                        // Get a match for all the columns in the row    
                        MatchCollection Columns = Regex.Matches(Row.Value, ColumnExpression, RegexOptions.Multiline | RegexOptions.Singleline | RegexOptions.IgnoreCase);

                        // Loop through each column element    
                        foreach (Match Column in Columns)
                        {

                            DataColumnCollection columns = dt.Columns;

                            if (!columns.Contains("Column " + iCurrentColumn))
                            {
                                //Add Columns  
                                dt.Columns.Add("Column " + iCurrentColumn);
                            }
                            // Add the value to the DataRow    
                            dr[iCurrentColumn] = Column.Groups[1].ToString();
                            // Increase the current column    
                            iCurrentColumn += 1;

                        }

                        // Add the DataRow to the DataTable    
                        dt.Rows.Add(dr);

                    }

                    // Increase the current row counter    
                    iCurrentRow += 1;
                }


            }

            return (dt);
        }
    }
}

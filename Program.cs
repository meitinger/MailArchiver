/* Copyright (C) 2011-2012, Manuel Meitinger
 * 
 * This program is free software: you can redistribute it and/or modify
 * it under the terms of the GNU General Public License as published by
 * the Free Software Foundation, either version 2 of the License, or
 * (at your option) any later version.
 * 
 * This program is distributed in the hope that it will be useful,
 * but WITHOUT ANY WARRANTY; without even the implied warranty of
 * MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
 * GNU General Public License for more details.
 * 
 * You should have received a copy of the GNU General Public License
 * along with this program.  If not, see <http://www.gnu.org/licenses/>.
 */

using System;
using System.Collections.Generic;
using System.Configuration;
using System.Data;
using System.Data.SqlClient;
using System.DirectoryServices;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using System.Security.Principal;
using System.Text;
using Aufbauwerk.Tools.MailArchiver.Properties;

namespace Aufbauwerk.Tools.MailArchiver
{
    internal static class Program
    {
        private class Context : Configuration.IContext
        {
            private readonly HashSet<string> groupDNs;
            private readonly com.sun.mail.imap.IMAPFolder folder;
            private HashSet<SecurityIdentifier> groups;

            public Context(IEnumerable<string> memberOf, com.sun.mail.imap.IMAPFolder folder)
            {
                this.groupDNs = new HashSet<string>(memberOf);
                this.folder = folder;
            }

            public bool IsMemberOf(SecurityIdentifier sid)
            {
                if (this.groups == null)
                {
                    this.groups = new HashSet<SecurityIdentifier>();
                    using (var directorySearcher = new DirectorySearcher("", new string[] { "objectSid", "memberOf" }))
                    {
                        var newGroupsDNs = new HashSet<string>(this.groupDNs);
                        while (newGroupsDNs.Count() > 0)
                        {
                            directorySearcher.Filter = string.Format("(&(objectCategory=group)(|{0}))", string.Join(string.Empty, newGroupsDNs.Select(dn => string.Format("(distinguishedName={0})", Program.EscapeLdap(dn))).ToArray()));
                            newGroupsDNs.Clear();
                            foreach (var group in directorySearcher.FindAll().Cast<SearchResult>())
                            {
                                this.groups.Add(new SecurityIdentifier((byte[])group.Properties["objectSid"][0], 0));
                                newGroupsDNs.UnionWith(group.Properties["memberOf"].Cast<string>());
                            }
                            newGroupsDNs.ExceptWith(this.groupDNs);
                            this.groupDNs.UnionWith(newGroupsDNs);
                        }
                    }
                }
                return this.groups.Contains(sid);
            }

            public int GetQuotaUsage(string resource)
            {
                var res = this.folder.getQuota()[0].resources.Where(r => r.name == resource).Single();
                return (int)(res.usage * 100 / res.limit);
            }

            public string Path { get { return this.folder.getFullName(); } }
        }

        private static readonly DateTime Epoche = new DateTime(1970, 1, 1, 0, 0, 0, DateTimeKind.Utc);

        [DllImport("kernel32.dll", ExactSpelling = true)]
        private static extern IntPtr GetConsoleWindow();

        private class NetOutputStream : java.io.OutputStream
        {
            private readonly Stream baseStream;

            public NetOutputStream(Stream baseStream)
            {
                this.baseStream = baseStream;
            }

            public override void flush()
            {
                this.baseStream.Flush();
            }

            public override void close()
            {
                this.baseStream.Close();
            }

            public override void write(byte[] b, int off, int len)
            {
                this.baseStream.Write(b, off, len);
            }

            public override void write(byte[] b)
            {
                this.baseStream.Write(b, 0, b.Length);
            }

            public override void write(int i)
            {
                this.baseStream.WriteByte((byte)i);
            }
        }

        public static java.util.Date ToJavaDate(this DateTime date)
        {
            return new java.util.Date((date - Program.Epoche).Ticks / 10000);
        }

        public static DateTime toFrameworkDate(this java.util.Date date)
        {
            return new DateTime(date.getTime() * 10000 + Program.Epoche.Ticks);
        }

        public static string getText(this javax.mail.Part p)
        {
            try
            {
                // if the part is a text, return it right away
                if (p.isMimeType("text/*"))
                    return p.getContent() as string;

                // if the part contains multiple alternative representations, find the plain text or html
                if (p.isMimeType("multipart/alternative"))
                {
                    var mp = (javax.mail.Multipart)p.getContent();
                    var text = (string)null;
                    var notHtmlYet = true;
                    for (int i = 0; i < mp.getCount(); i++)
                    {
                        javax.mail.Part bp = mp.getBodyPart(i);
                        if (bp.isMimeType("text/plain"))
                        {
                            // return plain text right away, since it is the prefered content
                            var plainText = bp.getContent() as string;
                            if (plainText != null)
                                return plainText;
                        }
                        else if (bp.isMimeType("text/html"))
                        {
                            // if we haven't found a html text yet, retrieve this one
                            if (text == null || notHtmlYet)
                            {
                                var html = bp.getContent() as string;
                                if (html != null)
                                {
                                    text = html;
                                    notHtmlYet = false;
                                }
                            }
                        }
                        else
                        {
                            // retrieve the unknown part if nothing has been found yet
                            if (text == null)
                                text = bp.getText();
                        }
                    }

                    // return the text
                    return text;
                }

                // if we have a unknown multipart, return the first usable content
                if (p.isMimeType("multipart/*"))
                {
                    var mp = (javax.mail.Multipart)p.getContent();
                    for (int i = 0; i < mp.getCount(); i++)
                    {
                        var s = mp.getBodyPart(i).getText();
                        if (s != null)
                            return s;
                    }
                }
            }
            catch (java.io.UnsupportedEncodingException e)
            {
                // notify the user about the unsupported encoding
                Console.WriteLine("      Unsupported encoding: {0}.", e.getMessage());
            }

            // tough luck...
            return null;
        }

        private static string EscapeLdap(string s)
        {
            var result = new StringBuilder(s.Length * 2);
            for (int i = 0; i < s.Length; i++)
            {
                switch (s[i])
                {
                    case '\\':
                        result.Append(@"\5c");
                        break;
                    case '*':
                        result.Append(@"\2a");
                        break;
                    case '(':
                        result.Append(@"\28");
                        break;
                    case ')':
                        result.Append(@"\29");
                        break;
                    case '\0':
                        result.Append(@"\00");
                        break;
                    case '/':
                        result.Append(@"\2f");
                        break;
                    default:
                        result.Append(s[i]);
                        break;
                }
            }
            return result.ToString();
        }

        private static void Main(string[] args)
        {
            try
            {
                // get the account list and display the account manager if requested
                if (Settings.Default.Accounts == null)
                {
                    Settings.Default.Upgrade();
                    if (Settings.Default.Accounts == null)
                        Settings.Default.Accounts = new Configuration.Account[0];
                    Settings.Default.Save();
                }
                var accounts = new List<Configuration.Account>(Settings.Default.Accounts);
                if (args.Length == 1 && args[0].Equals("/passwords", StringComparison.OrdinalIgnoreCase))
                {
                    System.Windows.Forms.Application.EnableVisualStyles();
                    System.Windows.Forms.Application.SetCompatibleTextRenderingDefault(false);
                    if (new AccountManager(accounts).ShowDialog(System.Windows.Forms.NativeWindow.FromHandle(GetConsoleWindow())) == System.Windows.Forms.DialogResult.OK)
                    {
                        Settings.Default.Accounts = accounts.ToArray();
                        Settings.Default.Save();
                    }
                    return;
                }
                if (args.Length > 0)
                    throw new ArgumentException(string.Format("USAGE: {0} [/passwords]", Environment.GetCommandLineArgs()[0]));
                if (Settings.Default.Accounts.Length == 0)
                    throw new ConfigurationErrorsException("No accounts are specified. Run the program again using the /passwords parameter.");

                // check the settings
                if (Settings.Default.Database == null)
                    throw new ConfigurationErrorsException("Database not specified");
                if (Settings.Default.Host == null)
                    throw new ConfigurationErrorsException("Host not specified");
                if (Settings.Default.Filter == null)
                    throw new ConfigurationErrorsException("Filter not specified");

                // prepare the imap properties
                var props = new java.util.Properties();
                string protocol;
                switch (Settings.Default.Security)
                {
                    case Configuration.SecurityLevel.None:
                        protocol = "imap";
                        break;
                    case Configuration.SecurityLevel.Ssl:
                        protocol = "imaps";
                        break;
                    case Configuration.SecurityLevel.Tls:
                        props.setProperty("mail.imap.starttls.enable", "true");
                        props.setProperty("mail.imap.starttls.required", "true");
                        protocol = "imap";
                        break;
                    default:
                        throw new NotImplementedException();
                }

                // connect to the domain and database
                using (var directorySearcher = new DirectorySearcher("", new string[] { "objectSid", "sAMAccountName", "userPrincipalName", "memberOf" }))
                using (var connection = new SqlConnection(Settings.Default.Database))
                {
                    connection.Open();
                    using (var deleteCommand = new SqlCommand(
                        "DELETE FROM [dbo].[Messages] " +
                        "WHERE [ID] = @ID AND [Owner] = @Owner", connection))
                    using (var insertCommand = new SqlCommand(
                        "INSERT INTO [dbo].[Messages] ([ID],[Mailbox],[Owner],[Date],[Sender],[Subject],[Body],[Data]) " +
                        "VALUES (@Id,@Mailbox,@Owner,@Date,@Sender,@Subject,@Body,@Data)", connection))
                    {
                        // create the delete parameter and prepare the command
                        var deleteMessageId = deleteCommand.Parameters.Add("ID", SqlDbType.VarChar, 250);
                        var deleteUserSid = deleteCommand.Parameters.Add("Owner", SqlDbType.VarBinary, 85);
                        deleteCommand.Prepare();

                        // create all insert parameters and prepare the command
                        var messageId = insertCommand.Parameters.Add("ID", SqlDbType.VarChar, 250);
                        var mailboxName = insertCommand.Parameters.Add("Mailbox", SqlDbType.NVarChar, -1);
                        var userSid = insertCommand.Parameters.Add("Owner", SqlDbType.VarBinary, 85);
                        var date = insertCommand.Parameters.Add("Date", SqlDbType.DateTime);
                        var sender = insertCommand.Parameters.Add("Sender", SqlDbType.NVarChar, -1);
                        var subject = insertCommand.Parameters.Add("Subject", SqlDbType.NVarChar, -1);
                        var body = insertCommand.Parameters.Add("Body", SqlDbType.NVarChar, -1);
                        var data = insertCommand.Parameters.Add("Data", SqlDbType.VarBinary, -1);
                        insertCommand.Prepare();

                        // go over all accounts
                        foreach (var account in accounts)
                        {
                            Console.WriteLine("{0}", account.Address);

                            // retrieve the user information
                            directorySearcher.Filter = string.Format("(&(objectCategory=user)(mail={0}))", Program.EscapeLdap(account.Address));
                            var userInfo = directorySearcher.FindOne();
                            if (userInfo == null)
                                throw new KeyNotFoundException();
                            userSid.Value = deleteUserSid.Value = (byte[])userInfo.Properties["objectSid"][0];
                            string userName;
                            switch (Settings.Default.Identity)
                            {
                                case Configuration.IdentityType.LocalPart:
                                    userName = account.User;
                                    break;
                                case Configuration.IdentityType.EmailAddress:
                                    userName = account.Address;
                                    break;
                                case Configuration.IdentityType.SamAccountName:
                                    userName = (string)userInfo.Properties["sAMAccountName"][0];
                                    break;
                                case Configuration.IdentityType.UserPrincipalName:
                                    userName = (string)userInfo.Properties["userPrincipalName"][0];
                                    break;
                                default:
                                    throw new NotImplementedException();
                            }

                            // connect to the imap server
                            var session = javax.mail.Session.getDefaultInstance(props);
                            var store = new com.sun.mail.imap.IMAPStore(session, new javax.mail.URLName(protocol, Settings.Default.Host, Settings.Default.Port, null, userName, account.Password));
                            store.connect();
                            try
                            {
                                // go over all mailboxes
                                foreach (var mailbox in store.getDefaultFolder().list("*").Cast<com.sun.mail.imap.IMAPFolder>())
                                {
                                    Console.WriteLine("  {0}", mailbox.getFullName());

                                    // retrieve the filter and open the mailbox
                                    var filter = Settings.Default.Filter.GetResult(new Context(userInfo.Properties["memberOf"].Cast<string>(), mailbox));
                                    if (filter.HasResult && !filter.Result)
                                        continue;
                                    mailboxName.Value = mailbox.getFullName();
                                    mailbox.open(javax.mail.Folder.READ_WRITE);
                                    try
                                    {
                                        // go over all messages
                                        foreach (var message in (filter.HasResult ? mailbox.getMessages() : mailbox.search(filter.SearchTerm)).Cast<com.sun.mail.imap.IMAPMessage>())
                                        {
                                            Console.WriteLine("    {0}", message.getMessageID());
                                            messageId.Value = deleteMessageId.Value = message.getMessageID();

                                            // get the sent date
                                            var received = message.getReceivedDate();
                                            date.Value = received == null ? (object)DBNull.Value : received.toFrameworkDate();

                                            // either get the sender address or a combination of all from addresses (usually only one)
                                            var senderAddress = message.getSender();
                                            if (senderAddress == null)
                                            {
                                                var fromAddresses = message.getFrom();
                                                sender.Value = fromAddresses == null ? (object)DBNull.Value : string.Join("; ", Array.ConvertAll(fromAddresses, address => address.toString()));
                                            }
                                            else
                                                sender.Value = senderAddress.toString();

                                            // get the subject
                                            subject.Value = message.getSubject() ?? (object)DBNull.Value;

                                            // get the plain text or html text
                                            body.Value = message.getText() ?? (object)DBNull.Value;

                                            // get the entire message
                                            using (var buffer = new MemoryStream())
                                            {
                                                message.writeTo(new NetOutputStream(buffer));
                                                data.Value = buffer.ToArray();
                                            }

                                            // move the message to the database
                                            using (var transaction = connection.BeginTransaction())
                                            {
                                                // delete any previously stored version of the message
                                                deleteCommand.Transaction = transaction;
                                                deleteCommand.ExecuteNonQuery();

                                                // insert this one, mark it as deleted and commit the db changes
                                                insertCommand.Transaction = transaction;
                                                insertCommand.ExecuteNonQuery();
                                                message.setFlag(javax.mail.Flags.Flag.DELETED, true);
                                                transaction.Commit();
                                            }
                                        }
                                    }
                                    finally { mailbox.close(true); }
                                }
                            }
                            finally { store.close(); }
                        }
                    }
                }
            }
            catch (Exception e)
            {
                // print the error (message) and set the exit code
#if DEBUG
                Console.Error.WriteLine(e);
#else
                Console.Error.WriteLine(e.Message);
#endif
                Environment.ExitCode = Marshal.GetHRForException(e);
            }
        }
    }
}

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
using System.ComponentModel;
using System.Linq;
using System.Net.Mail;
using System.Security.Cryptography;
using System.Security.Principal;
using System.Text;
using System.Text.RegularExpressions;
using System.Xml;
using System.Xml.Serialization;
using javax.mail.search;

namespace Aufbauwerk.Tools.MailArchiver.Configuration
{
    public class Account : INotifyPropertyChanged
    {
        private MailAddress mail;
        private string password;

        [XmlAttribute("EmailAddress")]
        public string Address
        {
            get { return this.mail == null ? string.Empty : this.mail.Address; }
            set
            {
                if (string.IsNullOrEmpty(value))
                    throw new ArgumentNullException("Address");
                if (this.Address == value)
                    return;
                var oldUser = this.User;
                var oldHost = this.Host;
                this.mail = new MailAddress(value);
                var newUser = this.User;
                var newHost = this.Host;
                this.OnPropertyChanged("Address");
                if (oldUser != newUser)
                    this.OnPropertyChanged("User");
                if (oldUser != newHost)
                    this.OnPropertyChanged("Host");
            }
        }

        [XmlIgnore]
        public string User { get { return this.mail == null ? string.Empty : this.mail.User; } }

        [XmlIgnore]
        public string Host { get { return this.mail == null ? string.Empty : this.mail.Host; } }

        [XmlIgnore]
        public string Password
        {
            get { return this.password ?? string.Empty; }
            set
            {
                if (string.IsNullOrEmpty(value))
                    throw new ArgumentNullException("Password");
                if (this.password == value)
                    return;
                this.password = value;
                this.OnPropertyChanged("Password");
            }
        }

        [XmlAttribute("Password")]
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        public byte[] EncryptedPassword
        {
            get { return ProtectedData.Protect(Encoding.UTF8.GetBytes(this.Password), null, DataProtectionScope.CurrentUser); }
            set { this.Password = Encoding.UTF8.GetString(ProtectedData.Unprotect(value, null, DataProtectionScope.CurrentUser)); }
        }

        protected virtual void OnPropertyChanged(string propertyName)
        {
            if (string.IsNullOrEmpty(propertyName))
                throw new ArgumentNullException("propertyName");
            if (this.PropertyChanged != null)
                this.PropertyChanged(this, new PropertyChangedEventArgs(propertyName));
        }

        public event PropertyChangedEventHandler PropertyChanged;
    }

    public enum SecurityLevel
    {
        None,
        Ssl,
        Tls,
    }

    public enum IdentityType
    {
        LocalPart,
        EmailAddress,
        UserPrincipalName,
        SamAccountName,
    }

    public class QueryResult
    {
        private SearchTerm searchTerm;
        private bool result;

        public QueryResult(bool result)
        {
            this.searchTerm = null;
            this.result = result;
        }

        public QueryResult(SearchTerm searchTerm)
        {
            if (searchTerm == null)
                throw new ArgumentNullException("searchTerm");
            this.searchTerm = searchTerm;
            this.result = false;
        }

        public static implicit operator QueryResult(bool result)
        {
            return new QueryResult(result);
        }

        public static implicit operator QueryResult(SearchTerm searchTerm)
        {
            return new QueryResult(searchTerm);
        }

        public bool HasResult { get { return this.searchTerm == null; } }

        public SearchTerm SearchTerm
        {
            get
            {
                if (this.HasResult) throw new InvalidOperationException();
                return this.searchTerm;
            }
        }

        public bool Result
        {
            get
            {
                if (!this.HasResult) throw new InvalidOperationException();
                return this.result;
            }
        }
    }

    public interface IContext
    {
        bool IsMemberOf(SecurityIdentifier sid);
        int GetQuotaUsage(string resource);
        string Path { get; }
    }

    public interface IQuery
    {
        QueryResult GetResult(IContext context);
    }

    public abstract class SingleQuery : IQuery
    {
        [XmlElement("And", typeof(And))]
        [XmlElement("Body", typeof(Body))]
        [XmlElement("Flag", typeof(Flag))]
        [XmlElement("Folder", typeof(Folder))]
        [XmlElement("From", typeof(From))]
        [XmlElement("Header", typeof(Header))]
        [XmlElement("MemberOf", typeof(MemberOf))]
        [XmlElement("MessageId", typeof(MessageId))]
        [XmlElement("MessageNumber", typeof(MessageNumber))]
        [XmlElement("Not", typeof(Not))]
        [XmlElement("Or", typeof(Or))]
        [XmlElement("Usage", typeof(Usage))]
        [XmlElement("ReceivedDate", typeof(ReceivedDate))]
        [XmlElement("Recipient", typeof(Recipient))]
        [XmlElement("SentDate", typeof(SentDate))]
        [XmlElement("Size", typeof(Size))]
        [XmlElement("Subject", typeof(Subject))]
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        public object Value { get; set; }

        [XmlIgnore]
        public IQuery SubQuery { get { return (IQuery)Value; } }

        public abstract QueryResult GetResult(IContext context);
    }

    public abstract class MultiQuery : IQuery
    {
        [XmlElement("And", typeof(And))]
        [XmlElement("Body", typeof(Body))]
        [XmlElement("Flag", typeof(Flag))]
        [XmlElement("Folder", typeof(Folder))]
        [XmlElement("From", typeof(From))]
        [XmlElement("Header", typeof(Header))]
        [XmlElement("MemberOf", typeof(MemberOf))]
        [XmlElement("MessageId", typeof(MessageId))]
        [XmlElement("MessageNumber", typeof(MessageNumber))]
        [XmlElement("Not", typeof(Not))]
        [XmlElement("Or", typeof(Or))]
        [XmlElement("Usage", typeof(Usage))]
        [XmlElement("ReceivedDate", typeof(ReceivedDate))]
        [XmlElement("Recipient", typeof(Recipient))]
        [XmlElement("SentDate", typeof(SentDate))]
        [XmlElement("Size", typeof(Size))]
        [XmlElement("Subject", typeof(Subject))]
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        public object[] Values { get; set; }

        [XmlIgnore]
        public IEnumerable<IQuery> SubQueries { get { return Values.Cast<IQuery>(); } }

        public abstract QueryResult GetResult(IContext context);
    }

    public class Filter : SingleQuery
    {
        public override QueryResult GetResult(IContext context)
        {
            return this.SubQuery.GetResult(context);
        }
    }

    public class And : MultiQuery
    {
        public override QueryResult GetResult(IContext context)
        {
            var terms = new List<SearchTerm>();
            foreach (var query in this.SubQueries)
            {
                var result = query.GetResult(context);
                if (result.HasResult)
                {
                    if (!result.Result)
                        return false;
                    continue;
                }
                terms.Add(result.SearchTerm);
            }
            switch (terms.Count)
            {
                case 0:
                    return true;
                case 1:
                    return terms[0];
                default:
                    return new AndTerm(terms.ToArray());
            }
        }
    }

    public class Body : IQuery
    {
        [XmlText]
        public string Pattern { get; set; }

        public QueryResult GetResult(IContext context)
        {
            return new BodyTerm(this.Pattern);
        }
    }

    public enum FlagName
    {
        Answered,
        Deleted,
        Draft,
        Flagged,
        Recent,
        Seen,
    }

    public class Flag : IQuery
    {
        [XmlText]
        public FlagName Name { get; set; }

        public QueryResult GetResult(IContext context)
        {
            javax.mail.Flags.Flag name;

            switch (this.Name)
            {
                case FlagName.Answered:
                    name = javax.mail.Flags.Flag.ANSWERED;
                    break;
                case FlagName.Deleted:
                    name = javax.mail.Flags.Flag.DELETED;
                    break;
                case FlagName.Draft:
                    name = javax.mail.Flags.Flag.DRAFT;
                    break;
                case FlagName.Flagged:
                    name = javax.mail.Flags.Flag.FLAGGED;
                    break;
                case FlagName.Recent:
                    name = javax.mail.Flags.Flag.RECENT;
                    break;
                case FlagName.Seen:
                    name = javax.mail.Flags.Flag.SEEN;
                    break;
                default:
                    throw new NotImplementedException();
            }
            return new FlagTerm(new javax.mail.Flags(name), true);
        }
    }

    public class Folder : IQuery
    {
        [XmlText]
        public string Path { get; set; }

        public QueryResult GetResult(IContext context)
        {
            return this.Path == context.Path;
        }
    }

    public class From : IQuery
    {
        [XmlText]
        public string Pattern { get; set; }

        public QueryResult GetResult(IContext context)
        {
            return new FromStringTerm(this.Pattern);
        }
    }

    public class Header : IQuery
    {
        [XmlAttribute]
        public string Name { get; set; }

        [XmlText]
        public string Pattern { get; set; }

        public QueryResult GetResult(IContext context)
        {
            return new HeaderTerm(this.Name, this.Pattern);
        }
    }

    public class MemberOf : IQuery
    {
        private readonly static Regex SidRegex = new Regex(@"^S-1-5(-\d+)+$", RegexOptions.ExplicitCapture);

        [XmlText]
        public string Group { get; set; }

        public QueryResult GetResult(IContext context)
        {
            return context.IsMemberOf(MemberOf.SidRegex.IsMatch(this.Group) ? new SecurityIdentifier(this.Group) : (SecurityIdentifier)new NTAccount(this.Group).Translate(typeof(SecurityIdentifier)));
        }
    }

    public class MessageId : IQuery
    {
        [XmlText]
        public string Value { get; set; }

        public QueryResult GetResult(IContext context)
        {
            return new MessageIDTerm(this.Value);
        }
    }

    public class MessageNumber : IQuery
    {
        [XmlText]
        public int Value { get; set; }

        public QueryResult GetResult(IContext context)
        {
            return new MessageNumberTerm(this.Value);
        }
    }

    public class Not : SingleQuery
    {
        public override QueryResult GetResult(IContext context)
        {
            var result = this.SubQuery.GetResult(context);
            if (result.HasResult)
                return !result.Result;
            return new NotTerm(result.SearchTerm);
        }
    }

    public class Or : MultiQuery
    {
        public override QueryResult GetResult(IContext context)
        {
            var terms = new List<SearchTerm>();
            foreach (var query in this.SubQueries)
            {
                var result = query.GetResult(context);
                if (result.HasResult)
                {
                    if (result.Result)
                        return true;
                    continue;
                }
                terms.Add(result.SearchTerm);
            }
            switch (terms.Count)
            {
                case 0:
                    return false;
                case 1:
                    return terms[0];
                default:
                    return new OrTerm(terms.ToArray());
            }
        }
    }

    public enum UsageType
    {
        DiskSpace,
        FileCount,
    }

    public class Usage : IQuery
    {
        [XmlAttribute("Of")]
        public UsageType Type { get; set; }

        [XmlAttribute("Is")]
        public ComparisonType Comparison { get; set; }

        [XmlText]
        public int Value { get; set; }

        public QueryResult GetResult(IContext context)
        {
            string type;
            switch (this.Type)
            {
                case UsageType.DiskSpace:
                    type = "STORAGE";
                    break;
                case UsageType.FileCount:
                    type = "MESSAGE";
                    break;
                default:
                    throw new NotImplementedException();
            }
            var used = context.GetQuotaUsage(type);
            switch (this.Comparison)
            {
                case ComparisonType.Equal:
                    return used == this.Value;
                case ComparisonType.GreaterOrEqual:
                    return used >= this.Value;
                case ComparisonType.GreaterThan:
                    return used > this.Value;
                case ComparisonType.LessOrEqual:
                    return used <= this.Value;
                case ComparisonType.LessThan:
                    return used < this.Value;
                case ComparisonType.NotEqual:
                    return used != this.Value;
                default:
                    throw new NotImplementedException();
            }
        }
    }

    public enum ComparisonType : int
    {
        Equal = ComparisonTerm.EQ,
        GreaterOrEqual = ComparisonTerm.GE,
        GreaterThan = ComparisonTerm.GT,
        LessOrEqual = ComparisonTerm.LE,
        LessThan = ComparisonTerm.LT,
        NotEqual = ComparisonTerm.NE,
    }

    public abstract class Date : IQuery
    {
        private static readonly DateTime Now = DateTime.Now;

        [XmlAttribute("Is")]
        public ComparisonType Comparison { get; set; }

        [XmlIgnore]
        public TimeSpan Offset { get; set; }

        [XmlText(DataType = "duration")]
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        public string Value
        {
            get { return XmlConvert.ToString(this.Offset); }
            set { this.Offset = XmlConvert.ToTimeSpan(value); }
        }

        [XmlAttribute]
        public int Year { get; set; }
        [XmlIgnore]
        public bool YearSpecified { get; set; }

        [XmlAttribute]
        public int Month { get; set; }
        [XmlIgnore]
        public bool MonthSpecified { get; set; }

        [XmlAttribute]
        public int Day { get; set; }
        [XmlIgnore]
        public bool DaySpecified { get; set; }

        [XmlAttribute]
        public int Hour { get; set; }
        [XmlIgnore]
        public bool HourSpecified { get; set; }

        [XmlAttribute]
        public int Minute { get; set; }
        [XmlIgnore]
        public bool MinuteSpecified { get; set; }

        [XmlAttribute]
        public int Second { get; set; }
        [XmlIgnore]
        public bool SecondSpecified { get; set; }

        protected abstract SearchTerm CreateTerm(int comparison, java.util.Date date);

        public QueryResult GetResult(IContext context)
        {
            return this.CreateTerm
            (
                (int)this.Comparison,
                (
                    new DateTime
                    (
                        this.YearSpecified ? this.Year : Date.Now.Year,
                        this.MonthSpecified ? this.Month : (this.YearSpecified ? 1 : Date.Now.Month),
                        this.DaySpecified ? this.Day : (this.YearSpecified || this.MonthSpecified ? 1 : Date.Now.Day),
                        this.HourSpecified ? this.Hour : (this.YearSpecified || this.MonthSpecified || this.DaySpecified ? 0 : Date.Now.Hour),
                        this.MinuteSpecified ? this.Minute : (this.YearSpecified || this.MonthSpecified || this.DaySpecified || this.HourSpecified ? 0 : Date.Now.Minute),
                        this.SecondSpecified ? this.Second : (this.YearSpecified || this.MonthSpecified || this.DaySpecified || this.HourSpecified || this.MinuteSpecified ? 0 : Date.Now.Second)
                    ) +
                    this.Offset
                ).ToJavaDate()
            );
        }
    }

    public class ReceivedDate : Date
    {
        protected override SearchTerm CreateTerm(int comparison, java.util.Date date)
        {
            return new ReceivedDateTerm(comparison, date);
        }
    }

    public enum RecipientType
    {
        To,
        Cc,
        Bcc,
    }

    public class Recipient : IQuery
    {
        [XmlAttribute("In")]
        public RecipientType Type { get; set; }

        [XmlText]
        public string Pattern { get; set; }

        public QueryResult GetResult(IContext context)
        {
            javax.mail.Message.RecipientType type;
            switch (this.Type)
            {
                case RecipientType.To:
                    type = javax.mail.Message.RecipientType.TO;
                    break;
                case RecipientType.Cc:
                    type = javax.mail.Message.RecipientType.CC;
                    break;
                case RecipientType.Bcc:
                    type = javax.mail.Message.RecipientType.BCC;
                    break;
                default:
                    throw new NotImplementedException();
            }
            return new RecipientStringTerm(type, this.Pattern);
        }
    }

    public class SentDate : Date
    {
        protected override SearchTerm CreateTerm(int comparison, java.util.Date date)
        {
            return new SentDateTerm(comparison, date);
        }
    }

    public class Size : IQuery
    {
        [XmlAttribute("Is")]
        public ComparisonType Comparison { get; set; }

        [XmlText]
        public int Value { get; set; }

        public QueryResult GetResult(IContext context)
        {
            return new SizeTerm((int)this.Comparison, this.Value);
        }
    }

    public class Subject : IQuery
    {
        [XmlText]
        public string Pattern { get; set; }

        public QueryResult GetResult(IContext context)
        {
            return new SubjectTerm(this.Pattern);
        }
    }
}

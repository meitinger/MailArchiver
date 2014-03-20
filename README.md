Mail Archiver
=============


Description
-----------
This utility archives (i.e. moves) mails from an IMAP server to a SQL database.
It allows for very fine-grained filtering and handling of multiple configurable
IMAP accounts that are being mapped to domain users by mail address.


Requirements
------------
To achieve this the program relies on the Sun Java Mail library and IKVM. The
IKVM assemblies as well as the recompiled Java Mail assembly (please use the
module name `Oracle.Java.Mail`) must be put into the `lib` folder; to avoid any
licensing issues, these files are not included in this repository.


Configuration
-------------
All settings apart from the accounts themselves are stored in the app config
file `MailArchive.exe.config`.

- *Database*: A connection string to access the SQL database.
- *Host*: The host name or IP address of the IMAP server.
- *Port*: The IMAP server's port, defaults to `143`.
- *Security*: The security level to be used, either `None`, `Ssl` or `Tls`.
- *Identity*: Describes what part of the mail adress should be used for
              authentication
              - `LocalPart`: Everything before the `@`.
              - `EmailAddress` (default): The whole mail address.
              - `SamAccountName`: The domain user's account name associated with
                                  the given mail address.
              - `UserPrincipalName`: The UPN of the associated domain account.
- Filter: Describes what messages should be archived. See the section about
          [filtering](#Filters) below.

### Accounts
To configure the accounts run the program with the command line

    MailArchiver.exe /passwords

This will bring up a dialog with a list containing one column for addresses and
another one for passwords. The passwords are saved as a per-user setting and
are encrypted using [Windows Data Protection](http://msdn.microsoft.com/en-us/library/ms995355.aspx).

### Filters
Basically, all search terms supported by Java Mail can be used, including but
not limited to the date the message was received, by whom it was sent, how big
it is in size, the subject etc. Additionally, you can also check the quota
usage and the domain groups that the associated account is a member of.

An example: Let's assume that we want *Mail Archiver* to actually only do some
cleanup if the quota usage exceeds 80%. If that is the case and the user is an
ordinary employee, all mails older than a year or larger than 3MB shall be
archived. If the user is a manager, we are a bit more cautious and use an age
of 3 years and a size of 10MB. Also, the *SPAM* folder should not be cleared
because it's using a text matching algorithm.

The resulting filter is:

    <Filter>
      <And>
        <Usage Of="DiskSpace" Is="GreaterThan">80</Usage>
        <Not>
          <Folder>Spam</Folder>
        </Not>
        <Or>
          <And>
            <MemberOf>CONTOSO\Employees</MemberOf>
            <Or>
              <ReceivedDate Is="LessThan">-P1Y</ReceivedDate>
              <Size Is="GreaterThan">3145728</Size>
            </Or>
          </And>
          <And>
            <MemberOf>CONTOSO\Managers</MemberOf>
            <Or>
              <ReceivedDate Is="LessThan">-P3Y</ReceivedDate>
              <Size Is="GreaterThan">10485760</Size>
            </Or>
          </And>
        </Or>
      </And>
    </Filter>

As mentioned earlier, there are a lot more filters available. Please have a
look at `Configuration.cs` to see what they are and what their syntax is.


Database Schema
---------------
When archiving a mail, the program extracts the sender, date, subject and body
(preferably plain-text, otherwise html) and stores them together with the raw
message and the associated user's SID in a table called `Messages`, which must
have the following columns:

    ID varchar(250) NOT NULL
    Mailbox nvarchar(max) NOT NULL
    Owner varbinary(85) NOT NULL
    Date datetime NULL
    Sender nvarchar(max) NULL
    Subject nvarchar(max) NULL
    Body nvarchar(max) NULL
    Data varbinary(max) NOT NULL

`Mailbox` contains the folder path to the message, the combination of `ID` and
`Owner` is unique and should be (at least) indexed.

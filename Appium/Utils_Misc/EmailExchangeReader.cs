using Microsoft.Exchange.WebServices.Data;
using System;
using System.IO;
using System.Security.Cryptography;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading;

namespace UIAutomation.Utils_Misc
{
    class EmailExchangeReader
    {
        private static readonly log4net.ILog log = log4net.LogManager.GetLogger(typeof(EmailExchangeReader));

        #region initialize

        //For Encryption
        // 32 bytes long.  Using a 16 character string here gives us 32 bytes when converted to a byte array.
        private const string initVector = "pemgail9uzpgzl88";
        // This constant is used to determine the keysize of the encryption algorithm.
        private const int keysize = 256;
        // This is the passphrase and has to match encrytion code
        private const string passPhrase = "Automation";
        const int emailsToRead = 12; // how many mails are read to match the subject in the mails
        const int pollForMailTimeOut = 180; // timeout in seconds

        #endregion initilize

        // Initilizes the Exchange service. Will be used as the starting point for most of the methods
        public static ExchangeService InitService(string username, string encryptedPassword)
        {
            ExchangeService service = new ExchangeService(ExchangeVersion.Exchange2013_SP1);
            service.Credentials = new WebCredentials(username, encryptedPassword);
            service.AutodiscoverUrl(username, RedirectionUrlValidationCallback);
            return service;
        }

        // Retrieves the body of an email and returns the body as a string
        public static string GetUnreadMailBody(string subject)
        {
            // Initialize service
            ExchangeService service = InitService(Data.Get("config_mail:username"), Data.Get("config_mail:password"));
            int timeOutCounter = 1;

            // Loop through Inbox till timeout or till mail is received
            do
            {
                // Search folder
                Folder sent = Folder.Bind(service, WellKnownFolderName.SentItems);
                SearchFilter sf = new SearchFilter.SearchFilterCollection(LogicalOperator.And, new SearchFilter.IsEqualTo(EmailMessageSchema.IsRead, true));
                ItemView view = new ItemView(emailsToRead);
                FindItemsResults<Item> findResult = service.FindItems(WellKnownFolderName.SentItems, sf, view);
                // Get relevant text
                foreach (Item i in findResult)
                {
                    if (i.Subject == subject)
                    {
                        EmailMessage message = EmailMessage.Bind(service, new ItemId(i.Id.ToString()), new PropertySet(BasePropertySet.IdOnly, ItemSchema.Body));
                        message.IsRead = true;
                        message.Update(ConflictResolutionMode.AutoResolve);
                        return message.Body;
                    }
                }
                Thread.Sleep(1000); // Explicit wait for polling
                timeOutCounter++;
            } while (timeOutCounter < pollForMailTimeOut);
            throw new Exception("No email recieved in the set timeout of '" + pollForMailTimeOut + "' with the set subject.");
        }

        // Retrieves the body of an email by looking up both the subject and specificText in the body.Returns the body as a string
        public static string GetUnreadMailBody(string subject, string partialBodySnippets)
        {
            // Initialize service
            ExchangeService service = InitService(Data.Get("config_mail:username"), Data.Get("config_mail:password"));
            int timeOutCounter = 1;
            string[] bodySnippetsArray = partialBodySnippets.Split(';');
            int totalSnippets = bodySnippetsArray.Length;

            // Loop through Inbox till timeout or till mail is received
            do
            {
                // Search folder
                Folder inbox = Folder.Bind(service, WellKnownFolderName.SentItems);
                SearchFilter sf = new SearchFilter.SearchFilterCollection(LogicalOperator.And, new SearchFilter.IsEqualTo(EmailMessageSchema.IsRead, false));
                ItemView view = new ItemView(emailsToRead);
                FindItemsResults<Item> findResult = service.FindItems(WellKnownFolderName.Inbox, sf, view);
                // Get relevant text
                foreach (Item i in findResult)
                {
                    if (i.Subject == subject)
                    {
                        EmailMessage message = EmailMessage.Bind(service, new ItemId(i.Id.ToString()), new PropertySet(BasePropertySet.IdOnly, ItemSchema.Body));
                        int counter = 0;
                        foreach (var x in bodySnippetsArray)
                        {
                            if (message.Body.ToString().Contains(x)) counter++;
                        }
                        if (counter == totalSnippets)
                        {
                            message.IsRead = true;
                            message.Update(ConflictResolutionMode.AutoResolve);
                            return message.Body;
                        }
                    }
                }
                Thread.Sleep(1000); // Explicit wait for polling
                timeOutCounter++;
            } while (timeOutCounter < pollForMailTimeOut);
            throw new Exception($"No email recieved in the set timeout of [{pollForMailTimeOut}] with subject: [{subject}] and partialEmailBodySnippets: [{partialBodySnippets}]");
        }

        // Extract the link in an email. With the start of the search being where the firstBodyWorld variable value starts.
        public static string GetUrlFromEmail(string subject, string firstBodyWorld = "")
        {
            string emailBody = GetUnreadMailBody(subject);
            string decodedBody = System.Net.WebUtility.HtmlDecode(emailBody);

            return ExtractURL(decodedBody, firstBodyWorld);
        }

        /// <summary>
        /// Extracting URL from decodedEmailBody param 
        /// </summary>
        /// <param name="decodedEmailBody"></param>
        /// <param name="firstBodyWorld"></param>
        /// <returns></returns>
        public static string ExtractURL(string decodedEmailBody, string firstBodyWorld = "")
        {
            string pattern = @"http\w{0,1}?:\/\/(www\.)?[-a-zA-Z0-9@:%._\+~#=]{2,256}\.[a-z]{2,6}\b([-a-zA-Z0-9@:%_\+.~#?&//=]*)";
            Regex regex = new Regex(pattern);
            //string match = regex.Match(emailBody).Value;
            string match = regex.Match(decodedEmailBody, decodedEmailBody.IndexOf(firstBodyWorld)).Value;
            return match;
        }

        public static void MarkEmailsAsReadBySubject(string subject)
        {
            int count = 0;
            // Initialize service
            ExchangeService service = InitService(Data.Get("config_mail:username"), Data.Get("config_mail:password"));
            SearchFilter sf = new SearchFilter.SearchFilterCollection(LogicalOperator.And, new SearchFilter.IsEqualTo(EmailMessageSchema.IsRead, false));
            ItemView view = new ItemView(100);
            FindItemsResults<Item> findResult = service.FindItems(WellKnownFolderName.Inbox, sf, view);
            // Get relevant text
            foreach (Item i in findResult)
            {
                if (i.Subject == subject)
                {
                    EmailMessage message = EmailMessage.Bind(service, new ItemId(i.Id.ToString()), new PropertySet(BasePropertySet.IdOnly, ItemSchema.Body));
                    message.IsRead = true;
                    message.Update(ConflictResolutionMode.AutoResolve);
                    count++;
                }
            }
            if (count > 0) log.Debug(count + " messages were marked as Read");
        }

        #region supportingMethods

        private static bool RedirectionUrlValidationCallback(string redirectionUrl)
        {
            // The default for the validation callback is to reject the URL.
            bool result = false;
            Uri redirectionUri = new Uri(redirectionUrl);
            // Validate the contents of the redirection URL. In this simple validation
            // callback, the redirection URL is considered valid if it is using HTTPS
            // to encrypt the authentication credentials. 
            if (redirectionUri.Scheme == "https")
            {
                result = true;
            }
            return result;
        }

        // Decrypts encryted password with passPhras as Automation
        public static string DecryptString(string cipherText)
        {
            byte[] initVectorBytes = Encoding.ASCII.GetBytes(initVector);
            byte[] cipherTextBytes = Convert.FromBase64String(cipherText);
            PasswordDeriveBytes password = new PasswordDeriveBytes(passPhrase, null);
            byte[] keyBytes = password.GetBytes(keysize / 8);
            RijndaelManaged symmetricKey = new RijndaelManaged();
            symmetricKey.Mode = CipherMode.CBC;
            ICryptoTransform decryptor = symmetricKey.CreateDecryptor(keyBytes, initVectorBytes);
            MemoryStream memoryStream = new MemoryStream(cipherTextBytes);
            CryptoStream cryptoStream = new CryptoStream(memoryStream, decryptor, CryptoStreamMode.Read);
            byte[] plainTextBytes = new byte[cipherTextBytes.Length];
            int decryptedByteCount = cryptoStream.Read(plainTextBytes, 0, plainTextBytes.Length);
            memoryStream.Close();
            cryptoStream.Close();
            return Encoding.UTF8.GetString(plainTextBytes, 0, decryptedByteCount);
        }

        #endregion supportingMethods
    }
}
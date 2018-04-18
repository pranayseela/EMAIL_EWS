using System;
using Microsoft.Exchange.WebServices.Data;

namespace EMAIL_EWS
{
    class Program
    {
        static void Main(string[] args)
        {
            ExchangeService service = new ExchangeService(ExchangeVersion.Exchange2007_SP1);

            service.Credentials = new WebCredentials("user1@contoso.com", "password");

            service.TraceEnabled = true;
            service.TraceFlags = TraceFlags.All;

            service.AutodiscoverUrl("user1@contoso.com", RedirectionUrlValidationCallback);

            #region To receive the attachements and save to local folder
            //FolderId objFId = new FolderId(WellKnownFolderName.Inbox);
            //FolderView objFView = new FolderView(10);
            //var inboxResult = service.FindFolders(objFId, objFView);
            //foreach (var item in inboxResult)
            //{
            //    item.Load();
            //    if (item.DisplayName.ToLower() == "inbox_subfolder_name") //mailbox folder name
            //    {
            //        var itemList = item.FindItems(new ItemView(10000));
            //        foreach (var mItem in itemList)
            //        {
            //            mItem.Load();
            //            var attachments = mItem.Attachments;

            //            foreach (var aItem in attachments)
            //            {
            //                if (aItem is FileAttachment)
            //                {
            //                    var fAttachment = aItem as FileAttachment;
            //                    //fAttachment.Load("C:\\Attachments\\" + fAttachment.Name);

            //                    //~\TEER\TEER\static\DTOBO\
            //                    fAttachment.Load("file_location" + fAttachment.Name); //Saves the attachment files to the folder.
            //                }
            //            }
            //        }
            //    }

            //}
            #endregion


            EmailMessage email = new EmailMessage(service);

            email.ToRecipients.Add("user1@contoso.com");

            email.Subject = "EWS Managed API - Hello World";
            email.Body = new MessageBody("This is the first email I've sent by using the EWS Managed API");

            email.Send();
            
        }

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
    }
}

using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Reflection;// to use Missing.Value
using Outlook = Microsoft.Office.Interop.Outlook;

namespace Send_HTMLTemplate_via_OutlookMail
{
    public class Class1
    {
        public string SendMailWithHTMLTemplate(string subject , string[] to,string[] cc ,string htmlTemplate, string[] imagePaths)
        {
            try
            {
                // Create the Outlook application.
                Outlook.Application oApp = new Outlook.Application();

                // Get the NameSpace and Logon information.
                Outlook.NameSpace oNS = oApp.GetNamespace("mapi");

                // Log on by using a dialog box to choose the profile.
                oNS.Logon(Missing.Value, Missing.Value, true, true);

                // Alternate logon method that uses a specific profile.
                // TODO: If you use this logon method, 
                //  change the profile name to an appropriate value.
                //oNS.Logon("YourValidProfile", Missing.Value, false, true); 

                // Create a new mail item.
                Outlook.MailItem oMsg = (Outlook.MailItem)oApp.CreateItem(Outlook.OlItemType.olMailItem);

                // string imagePath = @"C:\Users\Sikha.P\1mb.jpg";
                //  string imgBase64String = "data:image/png;base64," + ImageBase64;
                //Console.WriteLine(imgBase64String);

                // Set the subject.
                oMsg.Subject = subject;//"Send Using OOM in C#";
                int iteration = 0;
                string[] imageContents= new string[imagePaths.Length];


                foreach (string imagePath in imagePaths)
                {
                   
                    //Add an attachment.
                    String attachmentDisplayName = "MyAttachment_" + iteration;
                    // Attach the file to be embedded
                    string imageSrc = imagePath;//@"C:\Users\Sikha.P\logo.png"; // Change path as needed
                    Outlook.Attachment oAttach = oMsg.Attachments.Add(imageSrc, Outlook.OlAttachmentType.olByValue, null, attachmentDisplayName);
                    string imageContentid = "image_" + iteration; // Content ID can be anything. It is referenced in the HTML body
                    oAttach.PropertyAccessor.SetProperty("http://schemas.microsoft.com/mapi/proptag/0x3712001E", imageContentid);
                    imageContents[iteration] = imageContentid;
                    iteration++;
                }




                ////Add an attachment.
                //String attachmentDisplayName = "MyAttachment";
                //// Attach the file to be embedded
                //string imageSrc = imagePath;//@"C:\Users\Sikha.P\logo.png"; // Change path as needed
                //Outlook.Attachment oAttach = oMsg.Attachments.Add(imageSrc, Outlook.OlAttachmentType.olByValue, null, attachmentDisplayName);
                //string imageContentid = "image"; // Content ID can be anything. It is referenced in the HTML body
                //oAttach.PropertyAccessor.SetProperty("http://schemas.microsoft.com/mapi/proptag/0x3712001E", imageContentid);




                // Set HTMLBody.
                String sHtml;
                sHtml = String.Format(htmlTemplate, imageContents);
                oMsg.HTMLBody = sHtml;
                // Add a recipient.
                Outlook.Recipients oRecips = (Outlook.Recipients)oMsg.Recipients;
                foreach (string mailid in to)
                {
                    Outlook.Recipient oRecip = (Outlook.Recipient)oRecips.Add(mailid);
                    oRecip.Type = (int)Outlook.OlMailRecipientType.olTo;
                    oRecip.Resolve();
                }
                foreach (string mailid in cc)
                {
                    Outlook.Recipient oRecip = (Outlook.Recipient)oRecips.Add(mailid);
                    oRecip.Type = (int)Outlook.OlMailRecipientType.olCC;
                    oRecip.Resolve();
                }





                // Send.
                oMsg.Send();

                // Log off.
                oNS.Logoff();

                // Clean up.
                //oRecip = null;
                oRecips = null;
                oMsg = null;
                oNS = null;
                oApp = null;
                return "Email has been sent";
            }

            // Simple error handling.
            catch (Exception e)
            {
                return "{0} Exception caught."+ e;
            }
        }

        protected static string GetBase64StringForImage(string imgPath)
        {
            byte[] imageBytes = System.IO.File.ReadAllBytes(imgPath);
            string base64String = Convert.ToBase64String(imageBytes);
            return base64String;
        }
    }
}

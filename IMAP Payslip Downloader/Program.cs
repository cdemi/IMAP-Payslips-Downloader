using S22.Imap;
using System;
using System.IO;
using System.Linq;

namespace IMAP_Payslip_Downloader
{
    class Program
    {
        static void Main(string[] args)
        {
            var dir = $"{Environment.GetFolderPath(Environment.SpecialFolder.Desktop)}\\Payslips\\";
            using (ImapClient imapClient = new ImapClient("imap.gmail.com", 993, "", "", ssl: true))
            {
                var payslipsMailbox = imapClient.GetMailboxInfo("Payslips");

                var messageIDs = imapClient.Search(SearchCondition.All(), "Payslips");
                var messages = imapClient.GetMessages(messageIDs, seen: false, mailbox: "Payslips");

                if (!Directory.Exists(dir))
                {
                    Directory.CreateDirectory(dir);
                }

                foreach (var message in messages)
                {
                    var fileName = $"{message.Date().Value.ToString("yyyy-MM")}.pdf";

                    if (!File.Exists($"{dir}{fileName}"))
                    {
                        var payslip = message.Attachments.Where(a => a.Name.EndsWith(".pdf")).FirstOrDefault();

                        if (payslip != null)
                        {
                            Console.WriteLine($"Downloading {fileName}");
                            using (Stream payslipFile = File.Create($"{dir}{fileName}"))
                            {
                                payslip.ContentStream.CopyTo(payslipFile);
                            }
                        }
                    }
                }
            }
        }
    }
}

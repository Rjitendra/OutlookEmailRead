using System;

namespace Outlook_Read
{
    class Program
    {
        static void Main(string[] args)
        {
            try
            {
                var mails = OutLookEmails.ReadMailItems();
                int i = 1;

                if (mails.Count > 0)
                {

                    foreach (var mail in mails)
                    {
                        Console.WriteLine("Mail No :" + i);
                        Console.WriteLine("Mail Recieved From :" + mail.EmailFrom);
                        Console.WriteLine("Mail Subject :" + mail.EmailSubject);
                        Console.WriteLine("Mail Body :" + mail.EmailBody);
                        Console.WriteLine("");
                        i = i + 1;
                    }
                }
                else
                {
                    Console.WriteLine("No New Messages found!!");
                }
            }
            catch (Exception ex)
            {

                throw ex;
            }
            Console.ReadKey();
        }
    }

}

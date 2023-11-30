using System;
using System.Net;
using System.Net.Mail;

class Program
{
    static void Main()
    {
        try
        {
            MailMessage mail = new MailMessage();
            mail.From = new MailAddress("mail");
            mail.To.Add("ksawerynowak@lyson.com.pl");
            mail.Subject = "Temat e-maila";
            mail.Body = "Treść e-maila";

            // Dodawanie załącznika
            string sciezkaDoZalacznika = @"test.txt";
            Attachment attachment = new Attachment(sciezkaDoZalacznika);
            mail.Attachments.Add(attachment);

            // Konfiguracja klienta SMTP
            SmtpClient smtpClient = new SmtpClient("smtp.office365.com");
            smtpClient.Port = 587;
            smtpClient.EnableSsl = true;
            smtpClient.Credentials = new NetworkCredential("mail", "haslo");

            smtpClient.Send(mail);
            Console.WriteLine("E-mail został wysłany!");
        }
        catch (Exception ex)
        {
            Console.WriteLine("Błąd podczas wysyłania e-maila: " + ex.Message);
        }
    }
}

using ClosedXML.Excel;
using OfficeOpenXml;
using System;
using System.Net;
using System.Net.Mail;

class Program
{
    static void Main()
    {
        ServicePointManager.SecurityProtocol = SecurityProtocolType.Tls12; // Wymuszanie TLS 1.2

        var adresyEmail = PobierzAdresyEmailZExcela("Test2.xlsx");

        foreach (var adres in adresyEmail)
        {
            WyslijEmail(adres);
        }
    }

    public static void WyslijEmail(string adresEmail)
    {
        try
        {
            MailMessage mail = new MailMessage();
            mail.From = new MailAddress("sebastiandzidek@lyson.com.pl");
            mail.To.Add(adresEmail);
            mail.Subject = "spam_tester";
            mail.Body = "spam_tester";

            SmtpClient smtpClient = new SmtpClient("smtp.office365.com");
            smtpClient.Port = 587;
            smtpClient.EnableSsl = true;
            smtpClient.Credentials = new NetworkCredential("sebastiandzidek@lyson.com.pl", "Caq!!02973#");

            smtpClient.Send(mail);
            //Console.WriteLine("E-mail wysłany do: " + adresEmail);
            ZalogujPomyslne(adresEmail);
            Console.WriteLine($"E-mail wysłany do: {adresEmail}");
        }
        catch (Exception ex)
        {
            ZalogujBlad(adresEmail, ex);
            Console.WriteLine($"Błąd podczas wysyłania e-maila do: {adresEmail}: {ex.Message}");
        }
    }
    public static List<string> PobierzAdresyEmailZExcela(string sciezkaDoPliku)
    {
        var adresyEmail = new List<string>();

        // Użycie ClosedXML do odczytu pliku Excel
        using (var workbook = new XLWorkbook(sciezkaDoPliku))
        {
            var ws = workbook.Worksheet(1); // Załaduj pierwszy arkusz

            int lastRow = ws.LastRowUsed().RowNumber();
            // Zakładamy, że adresy e-mail są w pierwszej kolumnie
            for (int row = 1; row <= lastRow; row++)
            {
                var cellValue = ws.Cell(row, 1).Value.ToString();
                if (!string.IsNullOrWhiteSpace(cellValue))
                {
                    adresyEmail.Add(cellValue);
                }
            }
        }

        return adresyEmail;
    }
    public static void ZalogujBlad(string adresEmail, Exception ex)
    {
        // Ścieżka do pliku logu
        string sciezkaDoLogu = "log_bledow.txt";

        // Komunikat błędu
        string komunikat = $"Błąd podczas wysyłania e-maila do\t{adresEmail}\t{ex.Message}{Environment.NewLine}";

        // Zapis do pliku z użyciem try-catch, aby zapobiec przerywaniu programu przez błędy IO
        try
        {
            File.AppendAllText(sciezkaDoLogu, komunikat);
        }
        catch (IOException ioEx)
        {
            Console.WriteLine("Wystąpił problem z zapisem logu błędów: " + ioEx.Message);
        }
    }
    public static void ZalogujPomyslne(string adresEmail)
    {
        // Ścieżka do pliku logu
        string sciezkaDoLogu = "log_wyslanych.txt";

        // Komunikat błędu
        string komunikat = $"Pomyślnie wysłano do\t{adresEmail}{Environment.NewLine}";

        // Zapis do pliku z użyciem try-catch, aby zapobiec przerywaniu programu przez błędy IO
        try
        {
            File.AppendAllText(sciezkaDoLogu, komunikat);
        }
        catch (IOException ioEx)
        {
            Console.WriteLine("Wystąpił problem z zapisem logu błędów: " + ioEx.Message);
        }
    }
}

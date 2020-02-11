using Server.Models;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Net.Mail;
using System.Web;
using Xceed.Words.NET;

namespace Server.Classes
{
    public class DocCreator
    {
        const string FILENAME = @"C:\Users\N7_SPECTR\source\repos\DocCreator\doc.docx";// путь на свой поменяй
        DocX doc = DocX.Create(FILENAME);

        public void CreateDoc(Order order)
        {
            doc.PageLayout.Orientation = Xceed.Words.NET.Orientation.Landscape;

            List<string> list1 = new List<string> {
                "Приложение N 1",
                "к  постановлению Правительства РФ от 26.12.2011 N 1137",
                "(в редакции постановления Правительства РФ от 19.08.2017 N 981)" + Environment.NewLine
            };

            List<string> list2 = new List<string> {
                "СЧЕТ-ФАКТУРА № ",
                "ИСПРАВЛЕНИЕ № "
            };

            List<string> list3 = new List<string> {
                "Продавец " + order.Executor,
                "Адрес " + order.ExecutorAddress,
                "ИНН/КПП продавца " + order.ExecutorINN,
                "Грузоотправитель ",
                "Грузополучатель ",
                "К платежно-расчетному документу № ",
                "Покупатель " + order.Customer,
                "Адрес " + order.CustomerAddress,
                "ИНН/КПП покупателя " + order.CustomerINN,
                "Валюта: наименование, код Российский рубль, 643",
                "Идентификатор государственного контракта, договора (соглашения) (при наличии) " + Environment.NewLine
            };

            List<string> tableHeaders = new List<string> { "Наименование товара (описание выполненных работ, оказанных услуг), имущественного права",
                "Код вида товара", "Единица измерения", "Количество (объем)", "Цена (тариф) за единицу измерения",
                "Стоимость товаров (работ, услуг), имущественных прав без налога - всего", "В том числе сумма акциза", "Налоговая ставка",
                "Сумма налога, предъявляемая покупателю", "Стоимость товаров (работ, услуг), имущественных прав с налогом - всего",
                "Страна происхождения товара", "Регистрационный номер таможенной декларации"};

            Formatting format = new Formatting();
            format.FontFamily = new Xceed.Words.NET.Font("Times New Roman");
            format.Size = 8;
            foreach (string s in list1)
            {
                InsertParagraph(s, Alignment.right, format, 0);
            }

            format = new Formatting();
            format.FontFamily = new Xceed.Words.NET.Font("Times New Roman");
            format.Size = 12;
            format.Bold = true;
            foreach (string s in list2)
            {
                InsertParagraph(s, Alignment.left, format, 1);
            }

            format = new Formatting();
            format.FontFamily = new Xceed.Words.NET.Font("Times New Roman");
            format.Size = 12;
            foreach (string s in list3)
            {
                InsertParagraph(s, Alignment.left, format, 1);
            }

            Table table = doc.AddTable(1 + order.Materials.Split('%').Count(), tableHeaders.Count);
            for (int i = 0; i < tableHeaders.Count; i++)
            {
                table.Rows[0].Cells[i].Paragraphs.First().Append(tableHeaders.ElementAt(i));
                table.Rows[0].Cells[i].Paragraphs.First().Alignment = Alignment.center;
            }
            for (int i = 1; i < 1 + order.Materials.Split('%').Count(); i++)
            {
                table.Rows[i].Cells[0].Paragraphs.First().Append(order.Materials.Split('%')[i]);
                table.Rows[i].Cells[2].Paragraphs.First().Append("м2");
                table.Rows[i].Cells[3].Paragraphs.First().Append(order.Squares.Split('%')[i]);
                table.Rows[i].Cells[4].Paragraphs.First().Append(order.Prices.Split('%')[i]);
                table.Rows[i].Cells[4].Paragraphs.First().Append(order.Prices.Split('%')[i]);
            }
            //table.AutoFit = AutoFit.ColumnWidth;
            doc.InsertTable(table);
            doc.Save();
            try
            {
                MailMessage mail = new MailMessage();
                SmtpClient SmtpServer = new SmtpClient("smtp.mail.ru");
                mail.From = new MailAddress("твоя почта как снизу (там тоже поменять)");
                mail.To.Add("dmitriy.kazancev.2012@mail.ru");
                mail.Subject = "Test Mail - 1";
                mail.Body = "mail with attachment";

                Attachment attachment = new Attachment("C:\\Users\\N7_SPECTR\\source\\repos\\DocCreator\\doc.docx"); // путь тоже поменяй
                mail.Attachments.Add(attachment);

                SmtpServer.Port = 587;
                SmtpServer.Credentials = new System.Net.NetworkCredential("почта", "пароль");
                SmtpServer.EnableSsl = true;

                SmtpServer.Send(mail);
            }
            catch (Exception ex) { }
        }

        private void InsertParagraph(string text, Alignment alignment, Formatting format, double spacing)
        {
            Paragraph par = doc.InsertParagraph(text, false, format);
            par.Alignment = alignment;
            par.SpacingAfter(spacing);
        }
    }
}
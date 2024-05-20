using Microsoft.Office.Interop.Word;
using System.Runtime.InteropServices;

namespace ConsoleWordPractice;

public class DocumentGenerator
{
    public string Name { get; } = "DocumentGenerator";

    public void Execute()
    {
        Application app = new Application();
        Document document = app.Documents.Add();

        try
        {
            AddHeader(document,
                @"ОБЩЕСТВО С ОГРАНИЧЕННОЙ ОТВЕТСТВЕННОСТЬЮ «*********»
                ********, Г. ********, ******** ОБЛАСТЬ ****, Д. -. ОФ.
                ИНН ******** /КПП ********, ОГРН ********
                Р/СЧ: , БАНК: ******** БАНК ПАО
                «» Г. ********, БИК: ********,
                К/СЧ: ***,
                ТЕЛ/ФАКС: +7 **** **** , E-MAIL:@MAIL.RU");

            AddSubHeader(document, "КОММЕРЧЕСКОЕ ПРЕДЛОЖЕНИЕ от 19.05.2024 г.\nООО «**************» направляем Вам на рассмотрение коммерческое предложение на строительство ангара размером 20х50м.");

            Table table = document.Tables.Add(document.Range(document.Content.End - 1), 15, 5);
            table.Borders.Enable = 1;
            FormatTable(table);

            FillTableWithData(table);

            MergeColumns(table);

            AddFooter(document, "Стоимость строительства ангара размером 20х50м с боковой высотой 6 м составит; *.*** ***** (**** ********** ***** ******) рублей 00 копеек\nСрок строительства: 50 рабочих дней.\nУсловия оплаты: Рассрочка платежа до 19.10.2024г");
            string fileName = Guid.NewGuid().ToString();
            //document.SaveAs2($@"");
            document.SaveAs(@$"C:\Users\Роман\Desktop\Examples\{fileName}.docx");

            app.Documents.Open(@$"C:\Users\Роман\Desktop\Examples\{fileName}.docx");
            Console.ReadKey();

            document.Close();
        }
        catch (Exception ex)
        {
            Console.WriteLine(ex.Message);
        }
        finally
        {
            app.Quit();
            Marshal.ReleaseComObject(app);
        }
    }

    static void AddHeader(Document document, string text)
    {
        var headerRange = document
            .Sections[1]
            .Range;

        headerRange.Text = text;
        headerRange.Font.Size = 10;
        headerRange.Font.Color = (WdColor)WdColorRGB(62, 93, 120);
        headerRange.ParagraphFormat.Alignment = WdParagraphAlignment.wdAlignParagraphCenter;
    }

    static void AddSubHeader(Document document, string text)
    {
        var subHeaderRange = document.Content.Paragraphs.Add().Range;

        subHeaderRange.Text = text;
        subHeaderRange.Font.Size = 12;
        subHeaderRange.Font.Color = WdColor.wdColorBlack;
        subHeaderRange.ParagraphFormat.Alignment = WdParagraphAlignment.wdAlignParagraphCenter;
        subHeaderRange.InsertParagraphAfter();
    }

    static void FormatTable(Table table)
    {
        table.Rows[1].Range.Font.Bold = 1;
        table.Rows[1].Range.Font.Size = 11;
        table.Range.Font.Color = WdColor.wdColorBlack;

        table.Rows[1].Range.ParagraphFormat.Alignment = WdParagraphAlignment.wdAlignParagraphCenter;

        foreach (Row row in table.Rows)
        {
            foreach (Cell cell in row.Cells)
            {
                cell.Range.Font.Name = "Times New Roman";
                cell.Range.Font.Size = 10;
                cell.VerticalAlignment = WdCellVerticalAlignment.wdCellAlignVerticalCenter;
                cell.Range.ParagraphFormat.Alignment = WdParagraphAlignment.wdAlignParagraphCenter;
            }
        }

        foreach (Row row in table.Rows)
        {
            Cell cell = row.Cells[2];
            cell.Range.ParagraphFormat.Alignment = WdParagraphAlignment.wdAlignParagraphLeft;
        }

        table.Columns[1].Width = 40;
        table.Columns[2].Width = 200;
        table.Columns[3].Width = 40;
        table.Columns[4].Width = 50;
        table.Columns[5].Width = 75;
    }

    static void FillTableWithData(Table table)
    {
        table.Cell(1, 1).Range.Text = "№ п/п";
        table.Cell(1, 2).Range.Text = "Наименование";
        table.Cell(1, 3).Range.Text = "ед. изм.";
        table.Cell(1, 4).Range.Text = "Кол-во";
        table.Cell(1, 5).Range.Text = "Сумма";

        table.Cell(2, 1).Range.Text = "1.";
        table.Cell(2, 2).Range.Text = "Строительство ангара размером 20х50м с боковой высотой 6м:\nКомплектация:";
        table.Cell(2, 3).Range.Text = "М2";
        table.Cell(2, 4).Range.Text = "1000";
        table.Cell(2, 5).Range.Text = "6 400 000,00";

        table.Cell(3, 1).Range.Text = "1.1";
        table.Cell(3, 2).Range.Text = "Фундамент буронабивной, глубина 1500мм";

        table.Cell(4, 1).Range.Text = "1.2";
        table.Cell(4, 2).Range.Text = "Стойки – труба ф159х4,5 с шагом 3,0м";

        table.Cell(5, 1).Range.Text = "1.3";
        table.Cell(5, 2).Range.Text = "Фермы стропильные (ФС-1) – профильные трубы:\n- 80х80х3мм\n- 50х50х3мм";

        table.Cell(6, 1).Range.Text = "1.4";
        table.Cell(6, 2).Range.Text = "Связи вертикальные – профильные трубы:\n- 60х60х3";

        table.Cell(7, 1).Range.Text = "1.5";
        table.Cell(7, 2).Range.Text = "Связи горизонтальные – профильные трубы:\n- 50х50х3";

        table.Cell(8, 1).Range.Text = "1.5";
        table.Cell(8, 2).Range.Text = "Закладные детали – пластины лист:\n- T4мм";

        table.Cell(9, 1).Range.Text = "1.4";
        table.Cell(9, 2).Range.Text = "Прогоны:\nПК-1 - труба профильная 60х40х2мм с шагом 1м.\nПС-1- труба профильная 40х40х2мм с шагом  1м.";

        table.Cell(10, 1).Range.Text = "1.5";
        table.Cell(10, 2).Range.Text = "Наружная обшивка: профилированный лист \n- кровля - НС-35х0,5мм\n- стены – МП-20х0,5мм ";

        table.Cell(11, 1).Range.Text = "1.6";
        table.Cell(11, 2).Range.Text = "Фасонные элементы:\n- конек 250х250 цвет синий \n- ветровые планки 150х150 цвет синий ";

        table.Cell(12, 1).Range.Text = "1.7";
        table.Cell(12, 2).Range.Text = "Саморезы – 5.5х25 ОЦ";

        table.Cell(13, 1).Range.Text = "1.8";
        table.Cell(13, 2).Range.Text = "Ворота раздвижные  H=5,0 L=6,0 -  2 шт.";

        table.Cell(14, 1).Range.Text = "1.9";
        table.Cell(14, 2).Range.Text = "Грунтовка поверхности – ГФ-021";

        table.Cell(15, 1).Range.Text = "Итого: ";
        table.Cell(15, 5).Range.Text = "6 400 000,00";
    }

    static void AddFooter(Document document, string text)
    {
        var footerRange = document.Range(document.Content.End - 1);

        footerRange.Text = text;
        footerRange.Font.Size = 12;
        footerRange.Font.Color = WdColor.wdColorBlack;

        footerRange.ParagraphFormat.Alignment = WdParagraphAlignment.wdAlignParagraphLeft;
        footerRange.InsertParagraphAfter();
    }

    static int WdColorRGB(int red, int green, int blue)
    {
        return (red & 0xFF) | ((green & 0xFF) << 8) | ((blue & 0xFF) << 16);
    }

    static void MergeColumns(Table table)
    {
        table.Cell(2, 3).Merge(table.Cell(14, 3));
        table.Cell(2, 4).Merge(table.Cell(14, 4));
        table.Cell(2, 5).Merge(table.Cell(14, 5));


        table.Cell(15, 1).Merge(table.Cell(15, 4));
        table.Cell(15, 1).Range.ParagraphFormat.Alignment = WdParagraphAlignment.wdAlignParagraphLeft;
    }
}

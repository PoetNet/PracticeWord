using Microsoft.Office.Interop.Word;

Application app = new Application();

Document document = app.Documents.Add(Visible:true);

var r = document.Range();
//r.Text = "Hello, Word";
//r.Bold = 20;
//r.Italic = 1;

Table t = document.Tables.Add(r, 5, 5);
t.Borders.Enable = 1;

foreach (Row row in t.Rows)
{
    foreach (Cell cell in row.Cells)
    {
        if (cell.RowIndex == 1)
        {
            cell.Range.Text = $"Колонка {cell.ColumnIndex}";
            cell.Range.Bold = 1;
            cell.Range.Font.Name = "Times New Roman";
            cell.Range.Font.Size = 16;

            cell.VerticalAlignment = WdCellVerticalAlignment.wdCellAlignVerticalCenter;
            cell.Range.ParagraphFormat.Alignment = WdParagraphAlignment.wdAlignParagraphCenter;
        }
        else
        {
            cell.Range.Text = $"Row {cell.RowIndex}\nColumn: {cell.ColumnIndex}";
        }
    }
}


string fileName = Guid.NewGuid().ToString();
document.SaveAs(@$"J:\Practice\trash\{fileName}.docx");
app.Documents.Open(@$"J:\Practice\trash\{fileName}.docx");
Console.ReadKey();

try
{
    document.Close();
    app.Quit();
}
catch(Exception ex)
{
    Console.WriteLine(ex.Message);
}


Console.ReadKey();
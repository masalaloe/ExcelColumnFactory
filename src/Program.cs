using HRASystem.Libraries.ExcelUtilities;
using HRASystem.Libraries.ExcelUtilities.Factory;
//using ExcelUtilities;

Console.WriteLine("Hello, World!");

using (var package = new OfficeOpenXml.ExcelPackage())
{
    var worksheet = package.Workbook.Worksheets.Add("Temp");

    var columnFactory = new ColumnsFactory()
        .Builder()
        .AddColumn(col => {
            col.Header(h =>
            {
                h.Caption("table_caption1");
                h.Font(f =>
                {
                    f.Size(12);
                });
            })
            .Width(12)
            .Format("dd-MM-yyyy");
        })
        .AddColumn(col => col.Width(13).AddHeader("table_caption2"))
        //.AddColumn(col => {
        //    col.Header(h => h.Caption("table_caption2").Font(f => f.Size(12).Bold(false)))
        //        .Width(13).Format("dd-MM-yyyy").Validation(v => v.ValidationFormula(new List<string> { "Ya", "Tidak" }));
        //})
        .Build();

    //var columnFactory2 = new ColumnFactory()
    //    .Builder()
    //    .AddMultiColumns(col => {
    //        col.Header(h => h.Caption("table_caption2").Font(f => f.Size(12).Bold(false)))
    //            .Width(13).Format("dd-MM-yyyy").Validation(v => v.ValidationFormula(new List<string> { "Ya", "Tidak" }));

    //        col.Header(h => h.Caption("table_caption2").Font(f => f.Size(12).Bold(false)))
    //            .Width(13).Format("dd-MM-yyyy").Validation(v => v.ValidationFormula(new List<string> { "Ya", "Tidak" }));
    //    })
    //    .Build();

    //var columns = new ColumnFactory()
    //{
    //    new Column()
    //};

    //fileContents = package.GetAsByteArray();
}
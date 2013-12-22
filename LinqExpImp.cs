using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.IO;
//using LinqToExcel;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using System.Xml.Linq;
using System.Xml;
using System.Reflection;

namespace eCom.Interfaces.Excel
{
    public interface IExpImp<T>
    {
        IEnumerable<T> Import(Stream file, string fileExtension);
        MemoryStream Export(IEnumerable<T> objs);
    }
    public class LinqExpImp<T> : IExpImp<T>
    {
        private const string TEMP_EXCELS_DIR = @"C:\temp\Interfaces\Excel";
        public IEnumerable<T> Import(Stream file, string fileExtension)
        {
            IEnumerable<T> ret = null;
            
            using (SpreadsheetDocument document = SpreadsheetDocument.Open(file, false))
            {
                var sheet = document.WorkbookPart.WorksheetParts.FirstOrDefault().Worksheet;
                ret = document.ToTable<T>(sheet);
            }

            /*var excel = new ExcelQueryFactory();
            var indianaCompanies = from c in excel.Worksheet<T>()
                                   select c;*/
            
            return ret;
        }

        public MemoryStream Export(IEnumerable<T> objs)
        {
            using (SpreadsheetDocument document = SpreadsheetDocument.Open(GetTempFilePath(), true))
            {

            }
            MemoryStream ms = new MemoryStream();

            XmlWriterSettings xmlWriterSettings = new XmlWriterSettings() { Encoding = Encoding.UTF8 };
            XmlWriter xmlWriter = XmlWriter.Create(ms, xmlWriterSettings);

            objs.ToExcelXml().Save(xmlWriter);   //.Save() adds the <xml /> header tag!
            xmlWriter.Close();      //Must close the writer to dump it's content its output (the memory stream)

            return ms;
        }

        private static string GetTempFilePath(string extension = "xlsx")
        {
            if (!Directory.Exists(TEMP_EXCELS_DIR))
            {
                var parts = TEMP_EXCELS_DIR.Split('\\');
                var path = "";
                foreach (var dir in parts)
                {
                    path += dir + '\\';
                    if (!Directory.Exists(path))
                        Directory.CreateDirectory(path);
                }
            }

            return TEMP_EXCELS_DIR + @"\ProductsWithoutPicture." + extension;
        }
    }



    public static class ExcelExtensions
    {
        public static IEnumerable<T> ToTable<T>(this SpreadsheetDocument document, Worksheet sheet)
        {
            //Worksheet worksheet;
            SharedStringTable sharedString = document.WorkbookPart.SharedStringTablePart.SharedStringTable;
            //Initialize the customer list.
            List<T> result = new List<T>();

            //LINQ query to get first row with column names.
            var header = sheet.GetFirstChild<SheetData>().Elements<Row>().FirstOrDefault(r => r.RowIndex == 1);
            var headerValues = GetRowStrings(header, sharedString).ToList();
            var fieldMap = new Dictionary<int, FieldInfo>();
            var fields = typeof(T).GetFields(System.Reflection.BindingFlags.GetProperty | System.Reflection.BindingFlags.Public);
            foreach (var f in fields)
            {
                var i = headerValues.IndexOf(f.Name);
                if ( i >= 0 )
                    fieldMap.Add(i, f);
            }

            IEnumerable<Row> dataRows = sheet.GetFirstChild<SheetData>().Elements<Row>().Where(r => r.RowIndex > 1);
            foreach (var row in dataRows)
            {
                IEnumerable<String> textValues = GetRowStrings(row, sharedString);

                //Check to verify the row contained data.
                if (textValues.Count() > 0)
                {
                    //Create an object and add it to the list.
                    var textArray = textValues.ToArray();
                    T obj = Activator.CreateInstance<T>();
                    for (var i = 0; i < Math.Min(textValues.Count(), fieldMap.Count); i++) {
                        var f = fieldMap[i];
                        f.SetValue(obj, Convert.ChangeType(textArray[i], f.FieldType));
                    }
                    result.Add(obj);
                }
                else
                    //If no cells, then you have reached the end of the table.
                    break;
            }
            return result;
        }

        private static IEnumerable<String> GetRowStrings(Row row, SharedStringTable sharedString)
        {
            //LINQ query to return the row's cell values.
            //Where clause filters out any cells that do not contain a value.
            //Select returns the value of a cell unless the cell contains
            //  a Shared String.
            //If the cell contains a Shared String, its value will be a 
            //  reference id which will be used to look up the value in the 
            //  Shared String table.
            return row.Descendants<Cell>()
                    .Where(c => c.CellValue != null)
                    .Select(cell =>
                      (cell.DataType != null
                        && cell.DataType.HasValue
                        && cell.DataType == CellValues.SharedString
                      ? sharedString.ChildElements[
                        int.Parse(cell.CellValue.InnerText)].InnerText
                      : cell.CellValue.InnerText));
        }

        public static XDocument ToExcelXml(this IEnumerable<object> rows)
        {
            return rows.ToExcelXml("Sheet1");
        }
        public static XDocument ToExcelXml<T>(this IEnumerable<T> rows)
        {
            return rows.ToExcelXml("Sheet1");
        }

        public static XDocument ToExcelXml<T>(this IEnumerable<T> rows, string sheetName)
        {
            sheetName = sheetName.Replace("/", "-");
            sheetName = sheetName.Replace("\\", "-");

            XNamespace mainNamespace = "urn:schemas-microsoft-com:office:spreadsheet";
            XNamespace o = "urn:schemas-microsoft-com:office:office";
            XNamespace x = "urn:schemas-microsoft-com:office:excel";
            XNamespace ss = "urn:schemas-microsoft-com:office:spreadsheet";
            XNamespace html = "http://www.w3.org/TR/REC-html40";

            XDocument xdoc = new XDocument(new XDeclaration("1.0", "utf-8", "yes"));

            var headerRow = from p in rows.First().GetType().GetProperties()
                            select new XElement(mainNamespace + "Cell",
                                new XElement(mainNamespace + "Data",
                                    new XAttribute(ss + "Type", "String"), p.Name)); //Generate header using reflection

            XElement workbook = new XElement(mainNamespace + "Workbook",
                new XAttribute(XNamespace.Xmlns + "html", html),
                new XAttribute(XName.Get("ss", "http://www.w3.org/2000/xmlns/"), ss),
                new XAttribute(XName.Get("o", "http://www.w3.org/2000/xmlns/"), o),
                new XAttribute(XName.Get("x", "http://www.w3.org/2000/xmlns/"), x),
                new XAttribute(XName.Get("xmlns", ""), mainNamespace),
                new XElement(o + "DocumentProperties",
                        new XAttribute(XName.Get("xmlns", ""), o),
                        new XElement(o + "Author", "Smartdesk Systems Ltd"),
                        new XElement(o + "LastAuthor", "Smartdesk Systems Ltd"),
                        new XElement(o + "Created", DateTime.Now.ToString())
                    ), //end document properties
                new XElement(x + "ExcelWorkbook",
                        new XAttribute(XName.Get("xmlns", ""), x),
                        new XElement(x + "WindowHeight", 12750),
                        new XElement(x + "WindowWidth", 24855),
                        new XElement(x + "WindowTopX", 240),
                        new XElement(x + "WindowTopY", 75),
                        new XElement(x + "ProtectStructure", "False"),
                        new XElement(x + "ProtectWindows", "False")
                    ), //end ExcelWorkbook
                new XElement(mainNamespace + "Styles",
                        new XElement(mainNamespace + "Style",
                            new XAttribute(ss + "ID", "Default"),
                            new XAttribute(ss + "Name", "Normal"),
                            new XElement(mainNamespace + "Alignment",
                                new XAttribute(ss + "Vertical", "Bottom")
                            ),
                            new XElement(mainNamespace + "Borders"),
                            new XElement(mainNamespace + "Font",
                                new XAttribute(ss + "FontName", "Calibri"),
                                new XAttribute(x + "Family", "Swiss"),
                                new XAttribute(ss + "Size", "11"),
                                new XAttribute(ss + "Color", "#000000")
                            ),
                            new XElement(mainNamespace + "Interior"),
                            new XElement(mainNamespace + "NumberFormat"),
                            new XElement(mainNamespace + "Protection")
                        ),
                        new XElement(mainNamespace + "Style",
                            new XAttribute(ss + "ID", "Header"),
                            new XElement(mainNamespace + "Font",
                                new XAttribute(ss + "FontName", "Calibri"),
                                new XAttribute(x + "Family", "Swiss"),
                                new XAttribute(ss + "Size", "11"),
                                new XAttribute(ss + "Color", "#000000"),
                                new XAttribute(ss + "Bold", "1")
                            )
                        )
                    ), // close styles
                    new XElement(mainNamespace + "Worksheet",
                        new XAttribute(ss + "Name", sheetName /* Sheet name */),
                        new XElement(mainNamespace + "Table",
                            new XAttribute(ss + "ExpandedColumnCount", headerRow.Count()),
                            new XAttribute(ss + "ExpandedRowCount", rows.Count() + 1),
                            new XAttribute(x + "FullColumns", 1),
                            new XAttribute(x + "FullRows", 1),
                            new XAttribute(ss + "DefaultRowHeight", 15),
                            new XElement(mainNamespace + "Column",
                                new XAttribute(ss + "Width", 81)
                            ),
                            new XElement(mainNamespace + "Row", new XAttribute(ss + "StyleID", "Header"), headerRow),
                            from contentRow in rows
                            select new XElement(mainNamespace + "Row",
                                new XAttribute(ss + "StyleID", "Default"),
                                    from p in contentRow.GetType().GetProperties()
                                    select new XElement(mainNamespace + "Cell",
                                         new XElement(mainNamespace + "Data", new XAttribute(ss + "Type", "String"), p.GetValue(contentRow, null))) /* Build cells using reflection */ )
                        ), //close table
                        new XElement(x + "WorksheetOptions",
                            new XAttribute(XName.Get("xmlns", ""), x),
                            new XElement(x + "PageSetup",
                                new XElement(x + "Header",
                                    new XAttribute(x + "Margin", "0.3")
                                ),
                                new XElement(x + "Footer",
                                    new XAttribute(x + "Margin", "0.3")
                                ),
                                new XElement(x + "PageMargins",
                                    new XAttribute(x + "Bottom", "0.75"),
                                    new XAttribute(x + "Left", "0.7"),
                                    new XAttribute(x + "Right", "0.7"),
                                    new XAttribute(x + "Top", "0.75")
                                )
                            ),
                            new XElement(x + "Print",
                                new XElement(x + "ValidPrinterInfo"),
                                new XElement(x + "HorizontalResolution", 600),
                                new XElement(x + "VerticalResolution", 600)
                            ),
                            new XElement(x + "Selected"),
                            new XElement(x + "Panes",
                                new XElement(x + "Pane",
                                    new XElement(x + "Number", 3),
                                    new XElement(x + "ActiveRow", 1),
                                    new XElement(x + "ActiveCol", 0)
                                )
                            ),
                            new XElement(x + "ProtectObjects", "False"),
                            new XElement(x + "ProtectScenarios", "False")
                        ) // close worksheet options
                    ) // close Worksheet
                );

            xdoc.Add(workbook);

            return xdoc;
        }
    }


















    public class ExcelTest
    {
        static void Main(string[] args)
        {
          //Declare variables to hold refernces to Excel objects.
          Workbook workBook;
          SharedStringTable sharedStrings;
          IEnumerable<Sheet> workSheets;
          WorksheetPart custSheet;
          WorksheetPart orderSheet;

          //Declare helper variables.
          string custID;
          string orderID;
          List<Customer> customers;
          List<Order> orders;

          //Open the Excel workbook.
          using (SpreadsheetDocument document =
            SpreadsheetDocument.Open(@"C:\Temp\LinqSample.xlsx", true))
          {
            //References to the workbook and Shared String Table.
            workBook = document.WorkbookPart.Workbook;
            workSheets = workBook.Descendants<Sheet>();
            sharedStrings =
              document.WorkbookPart.SharedStringTablePart.SharedStringTable;

            //Reference to Excel Worksheet with Customer data.
            custID =
              workSheets.First(s => s.Name == @"Customer").Id;
            custSheet =
              (WorksheetPart)document.WorkbookPart.GetPartById(custID);

            //Load customer data to business object.
            customers =
              Customer.LoadCustomers(custSheet.Worksheet, sharedStrings);

            //Reference to Excel worksheet with order data.
            orderID =
              workSheets.First(sheet => sheet.Name == @"Order").Id;
            orderSheet =
              (WorksheetPart)document.WorkbookPart.GetPartById(orderID);

            //Load order data to business object.
            orders =
              Order.LoadOrders(orderSheet.Worksheet, sharedStrings);

            //List all customers to the console.
            //Write header information to the console.
            Console.WriteLine("All Customers");
            Console.WriteLine("{0, -15} {1, -15} {2, -5}",
              "Customer", "City", "State");

            //LINQ Query for all customers.
            IEnumerable<Customer> allCustomers =
                from customer in customers
                select customer;

            //Execute query and write customer information to the console.
            foreach (Customer c in allCustomers)
            {
              Console.WriteLine("{0, -15} {1, -15} {2, -5}",
                c.Name, c.City, c.State);
            }
            Console.WriteLine();
            Console.WriteLine();


            //Write all orders over $100 to the console.
            //Write header information to the console.
            Console.WriteLine("All Orders over $100");
            Console.WriteLine("{0, -15} {1, -10} {2, 10} {3, -5}",
              "Customer", "Date", "Amount", "Status");

            //LINQ Query for all orders over $100.
            //Join used to display customer information for the order.
            var highOrders =
              from customer in customers
              join order in orders on customer.Name equals order.Customer
              where order.Amount > 100.00
              select new
              {
                customer.Name,
                order.Date,
                order.Amount,
                order.Status
              };

            //Execute query and write information to the console.
            foreach (var result in highOrders)
            {
              Console.WriteLine("{0, -15} {1, -10} {2, 10} {3, -5}",
                result.Name, result.Date.ToShortDateString(),
                result.Amount, result.Status);
            }
            Console.WriteLine();
            Console.WriteLine();


            //Report on customer orders by status.
            //Write header information to  the console.
            Console.WriteLine("Customer Orders by Status");

            //LINQ Query for summarizing customer order information by status.
            //There are two LINQ queries.  
            //Internal query is used to group orders together by status and 
            //calculates the total order amount and number of orders.
            //External query is used to join Customer information.
            var sumoforders =
              from customer in customers
              select new
              {
                customer.Name,
                statusTotals =
                    from order in orders
                    where order.Customer == customer.Name
                    group order.Amount by order.Status into statusGroup
                    select new
                    {
                      status = statusGroup.Key,
                      orderAmount = statusGroup.Sum(),
                      orderCount = statusGroup.Count()
                    }
              };

            //Execute query and write information to the console.
            foreach (var customer in sumoforders)
            {
              //Write Customer name to the console.
              Console.WriteLine("-{0}-", customer.Name);
              foreach (var x in customer.statusTotals)
              {
                Console.WriteLine("  {0, -10}: {2,2} orders totaling {1, 7}",
                  x.status, x.orderAmount, x.orderCount);
              }
              Console.WriteLine();
            }

            //Keep the console window open.
            Console.Read();
          }
        }
        }
    /// <summary>
    /// Used to store customer information for analysis.
    /// </summary>
    public class Customer
    {
        //Properties.
        public string Name { get; set; }
        public string City { get; set; }
        public string State { get; set; }

        /// <summary>
        /// Helper method for creating a list of customers 
        /// from an Excel worksheet.
        /// </summary>
        public static List<Customer> LoadCustomers(Worksheet worksheet,
        SharedStringTable sharedString)
        {
        //Initialize the customer list.
        List<Customer> result = new List<Customer>();

        //LINQ query to skip first row with column names.
        IEnumerable<Row> dataRows =
            from row in worksheet.Descendants<Row>()
            where row.RowIndex > 1
            select row;

        foreach (Row row in dataRows)
        {
            //LINQ query to return the row's cell values.
            //Where clause filters out any cells that do not contain a value.
            //Select returns the value of a cell unless the cell contains
            //  a Shared String.
            //If the cell contains a Shared String, its value will be a 
            //  reference id which will be used to look up the value in the 
            //  Shared String table.
            IEnumerable<String> textValues =
            from cell in row.Descendants<Cell>()
            where cell.CellValue != null
            select
                (cell.DataType != null
                && cell.DataType.HasValue
                && cell.DataType == CellValues.SharedString
                ? sharedString.ChildElements[
                int.Parse(cell.CellValue.InnerText)].InnerText
                : cell.CellValue.InnerText)
            ;

            //Check to verify the row contained data.
            if (textValues.Count() > 0)
            {
            //Create a customer and add it to the list.
            var textArray = textValues.ToArray();
            Customer customer = new Customer();
            customer.Name = textArray[0];
            customer.City = textArray[1];
            customer.State = textArray[2];
            result.Add(customer);
            }
            else
            {
            //If no cells, then you have reached the end of the table.
            break;
            }
        }

        //Return populated list of customers.
        return result;
        }
    }

    /// <summary>
    /// Used to store order information for analysis.
    /// </summary>
    public class Order
    {
        //Properties.
        public string Number { get; set; }
        public DateTime Date { get; set; }
        public string Customer { get; set; }
        public Double Amount { get; set; }
        public string Status { get; set; }

        /// <summary>
        /// Helper method for creating a list of orders 
        /// from an Excel worksheet.
        /// </summary>
        public static List<Order> LoadOrders(Worksheet worksheet,
        SharedStringTable sharedString)
        {
        //Initialize order list.
        List<Order> result = new List<Order>();

        //LINQ query to skip first row with column names.
        IEnumerable<Row> dataRows =
            from row in worksheet.Descendants<Row>()
            where row.RowIndex > 1
            select row;

        foreach (Row row in dataRows)
        {
            //LINQ query to return the row's cell values.
            //Where clause filters out any cells that do not contain a value.
            //Select returns cell's value unless the cell contains
            //  a shared string.
            //If the cell contains a shared string its value will be a 
            //  reference id which will be used to look up the value in the 
            //  shared string table.
            IEnumerable<String> textValues =
            from cell in row.Descendants<Cell>()
            where cell.CellValue != null
            select
                (cell.DataType != null
                && cell.DataType.HasValue
                && cell.DataType == CellValues.SharedString
                ? sharedString.ChildElements[
                int.Parse(cell.CellValue.InnerText)].InnerText
                : cell.CellValue.InnerText)
            ;

            //Check to verify the row contains data.
            if (textValues.Count() > 0)
            {
            //Create an Order and add it to the list.
            var textArray = textValues.ToArray();
            Order order = new Order();
            order.Number = textArray[0];
            order.Date = new DateTime(1900, 1, 1).AddDays(
                Double.Parse(textArray[1]) - 2);
            order.Customer = textArray[2];
            order.Amount = Double.Parse(textArray[3]);
            order.Status = textArray[4];
            result.Add(order);
            }
            else
            {
            //If no cells, then you have reached the end of the table.
            break;
            }
        }

        //Return populated list of orders.
        return result;
        }
    }
}

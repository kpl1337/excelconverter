using ClosedXML.Excel;
using DocumentFormat.OpenXml.Bibliography;
using DocumentFormat.OpenXml.Spreadsheet;
using ExcelConverterLibrary.Models;
using System.Data;
using System.Runtime.ConstrainedExecution;
using System.Runtime.Serialization;
using System.Xml.Schema;
using System.Xml.Serialization;

namespace ExcelConverterLibrary
{
    public class XmlExcelConverter : IExcelConverter
    {
        /// <summary>
        /// Method loads excel file and return converted json.
        /// </summary>
        /// <param name="filePath">filePath to excel</param>
        /// <returns>Serialized json to string</returns>
        public string Convert(string filePath)
        {
            // ziskame data jako u JSONu
            List<Client> clients = ConvertExcelToClients(filePath);

            XmlSerializer ser = new XmlSerializer(typeof(List<Client>));

            using (StringWriter writer = new StringWriter())
            {
                ser.Serialize(writer, clients);

                return writer.ToString();
            }
        }

        private static List<Client> ConvertExcelToClients(string filePath)
        {
            Dictionary<string, Client> clients = new Dictionary<string, Client>();

            using (var workbook = new XLWorkbook(filePath))
            {
                IXLWorksheet sheet = workbook.Worksheet(1);
                IXLRangeRows rows = sheet.RangeUsed().RowsUsed();

                foreach (var rangeRow in rows.Skip(1)) // Skip header
                {
                    ProcessRow(rangeRow, sheet, clients);
                }
            }

            return new List<Client>(clients.Values);
        }

        private static void ProcessRow(IXLRangeRow row, IXLWorksheet sheet, Dictionary<string, Client> clients)
        {
            // Fixed columns
            string clientName = row.Cell(1).GetString();
            string companyID = row.Cell(2).GetString();
            string orderName = row.Cell(3).GetString();

            if (string.IsNullOrWhiteSpace(clientName) || string.IsNullOrWhiteSpace(companyID) || string.IsNullOrWhiteSpace(orderName))
                return;

            if (!clients.ContainsKey(companyID))
            {
                clients[companyID] = new Client
                {
                    Name = clientName,
                    CompanyID = companyID,
                    Orders = new List<Order>()
                };
            }

            var order = new Order
            {
                OrderName = orderName,
                ProductionUnits = ProcessProductionUnits(row, sheet)
            };

            clients[companyID].Orders.Add(order);
        }

        private static List<ProductionUnit> ProcessProductionUnits(IXLRangeRow row, IXLWorksheet sheet)
        {
            var productionUnits = new List<ProductionUnit>();

            // dynamic columns           
            for (int col = 4; col <= sheet.LastColumnUsed().ColumnNumber(); col++)
            {
                // get value
                var cell = row.Cell(col);
                if (cell.IsEmpty()) continue; // return if no value for the period

                string period = sheet.Cell(1, col).GetDateTime().ToString("MM/yyyy", System.Globalization.CultureInfo.InvariantCulture);
                int quantity = cell.GetValue<int>();

                productionUnits.Add(new ProductionUnit
                {
                    Period = period,
                    Quantity = quantity
                });
            }

            return productionUnits;
        }
    }
}

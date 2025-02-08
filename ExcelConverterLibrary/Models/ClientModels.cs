using Newtonsoft.Json;

namespace ExcelConverterLibrary.Models
{
    public class Client
    {
        [JsonProperty("Název klienta")]
        public required string Name { get; set; }

        [JsonProperty("IČ klienta")]
        public string? CompanyID { get; set; }

        [JsonProperty("Zakázky")]
        public required List<Order> Orders { get; set; }
    }

    public class Order
    {
        [JsonProperty("Název zakázky")]
        public required string OrderName { get; set; }

        [JsonProperty("Vyrobené kusy")]
        public List<ProductionUnit>? ProductionUnits { get; set; }
    }

    public class ProductionUnit
    {
        [JsonProperty("Období")]
        public required string Period { get; set; }

        [JsonProperty("Počet kusů")]
        public int Quantity { get; set; }
    }
}


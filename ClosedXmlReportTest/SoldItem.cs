namespace ClosedXmlReportTest
{
    class SoldItem
    {
        public int ShowableId { get; set; }
        public string Name { get; set; }
        public int OrderCount { get; set; }
        public int Price { get; set; }
        public int TotalSale { get => OrderCount * Price; }

        public SoldItem(int showableId, string name, int orderCount, int price)
        {
            ShowableId = showableId;
            Name = name;
            OrderCount = orderCount;
            Price = price;
        }

        public override string ToString()
        {
            return "ShowableId { ShowableId=" + ShowableId + ", Name=" + Name
                + ", OrderCount=" + OrderCount + ", Price=" + Price
                + ", TotalSale=" + TotalSale + "}";
        }
    }
}

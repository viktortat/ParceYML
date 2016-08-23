using System.Drawing;

namespace ParceYmlApp
{
    class RowItem
    {
        public int ColNom { get; set; }
        public int RowNom { get; set; }
        public string ParentId { get; set; }
        public string Name { get; set; }
        public Color Color { get; set; }
        public string NameCol { get; set; }
        public string Unit { get; set; }
        public string InnerText { get; set; }
    }


    class RowItemParam
    {
        public long Row_id { get; set; }
        public long ParamId { get; set; }
        public string Name { get; set; }
        public string NameTbn { get; set; }
        public string ParamType { get; set; }
    }

    //Row_id Available   ProductId Name    ProdType ProdKind    Url Price   CurrencyId CategoryId 
    //CategoryName Delivery    VendorCode Vendor  Description Picture

    class RowItemProduct
    {
        public long Row_id { get; set; }
        public bool Available { get; set; }
        public string ProductId { get; set; }
        public string Name { get; set; }
        public string ProdType { get; set; }
        public string ProdKind { get; set; }
        public string Url { get; set; }
        public decimal Price { get; set; }
        public string CurrencyId { get; set; }
        public string CategoryId { get; set; }
        public string CategoryName { get; set; }
        public string Delivery { get; set; }
        public string VendorCode { get; set; }
        public string Vendor { get; set; }
        public string Description { get; set; }
        public string Picture { get; set; }
    }
}
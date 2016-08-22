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
}
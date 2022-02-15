namespace DesignerEyesService.Entities
{
    public class Order
    {
        public string ID { get; set; }
        public string OrderNumber { get; set; }
        public string EntryID { get; set; }
        public string Status { get; set; }
        public string ItemLookupCode { get; set; }
        public string SimpleProdLineNo { get; set; }
        public string isPendingForwarding { get; set; }
        public string wasForwarded { get; set; }
        public string RealQty { get; set; }
        public string QtyOrdered { get; set; }
        public string QtyCancelled { get; set; }
        public string QtyRefunded { get; set; }
        public string QtyShipped { get; set; }
        public string City { get; set; }
        public string Country { get; set; }
        public string Region { get; set; }
        public string PostCode { get; set; }
        public string Street { get; set; }
        public string Street2 { get; set; }
        public string Telephone { get; set; }
        public string ContactEmail { get; set; }
        public string FirstName { get; set; }
        public string LastName { get; set; }
        public string RegionID { get; set; }

    }
}

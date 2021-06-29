using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace HermesLabelCreator.Models
{
    public class Shipment
    {
        [DisplayName("ShippmentID")]
        public string ShipmentID { get; set; }
        [DisplayName("Hermes AT Barcode")]
        public string HermesAtBarcode { get; set; }
        [DisplayName("UserID")]
        public string UserID { get; set; }
        [DisplayName("PW")]
        public string Password { get; set; }
        [DisplayName("Client")]
        public string Client { get; set; }
        [DisplayName("ArticleQuantity")]
        public string ArticleQuantity { get; set; }
        [DisplayName("ArticleNumber")]
        public string ArticleNumber { get; set; }
        [DisplayName("ArticleDescription")]
        public string ArticleDescription { get; set; }
        [DisplayName("ArticleWeight")]
        public string ArticleWeight { get; set; }
        [DisplayName("PUNumber")]
        public string PuNumber { get; set; }
        [DisplayName("ShipperTitle")]
        public string ShipperTitle { get; set; }
        [DisplayName("ShipperFirstName")]
        public string ShipperFirstName { get; set; }
        [DisplayName("ShipperLastName")]
        public string ShipperLastName { get; set; }
        [DisplayName("ShipperStreet")]
        public string ShipperStreet { get; set; }
        [DisplayName("ShipperStreetNr")]
        public string ShipperStreetNumber { get; set; }
        [DisplayName("ShipperAddressAdd")]
        public string ShipperAddressAdd { get; set; }
        [DisplayName("ShipperZip")]
        public string ShipperZip { get; set; }
        [DisplayName("ShipperCity")]
        public string ShipperCity { get; set; }
        [DisplayName("ShipperMobile")]
        public string ShipperMobile { get; set; }
        [DisplayName("ShipperCountry")]
        public string ShipperCountry { get; set; }
        [DisplayName("ShipperEmail")]
        public string ShipperEmail { get; set; }
        [DisplayName("ReceiverTitle")]
        public string ReceiverTitle { get; set; }
        [DisplayName("ReceiverFirstName")]
        public string ReceiverFirstName { get; set; }
        [DisplayName("ReceiverLastName")]
        public string ReceiverLastName { get; set; }
        [DisplayName("ReceiverStreet")]
        public string ReceiverStreet { get; set; }
        [DisplayName("ReceiverStreetNr")]
        public string ReceiverStreetNumber { get; set; }
        [DisplayName("ReceiverAddressAdd")]
        public string ReceiverAddressAdd { get; set; }
        [DisplayName("ReceiverZip")]
        public string ReceiverZip { get; set; }
        [DisplayName("RecevierCity")]
        public string ReceiverCity { get; set; }
        [DisplayName("ReceiverCountry")]
        public string ReceiverCountry { get; set; }
        [DisplayName("ReceiverMobile")]
        public string ReceiverMobile { get; set; }
        [DisplayName("ReceiverEmail")]
        public string ReceiverEmail { get; set; }
        [DisplayName("NextDay")]
        public string NextDay { get; set; }
        [DisplayName("Sperrgut")]
        public string BulkyGoods { get; set; }
        [DisplayName("KWS")]
        public string Kws { get; set; }
        //[DisplayName("2ManHandling")]
        //public string TwoManHandling { get; set; }
        //[DisplayName("Colli")]
        //public string Colli { get; set; }
        [DisplayName("Retour")]
        public string Returns { get; set; }
        [DisplayName("Date")]
        public string Date { get; set; }
        [DisplayName("OrderNumber")]
        public string OrderNumber { get; set; }
    }
}

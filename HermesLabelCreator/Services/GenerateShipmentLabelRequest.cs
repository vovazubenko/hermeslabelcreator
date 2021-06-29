using Newtonsoft.Json;
using System.Collections.Generic;

namespace HermesLabelCreator.Services
{
    public class GenerateShipmentLabelRequest
    {
        public DeliveryInfo Delivery { get; set; }
        public ReturnsInfo Returns { get; set; }

        public class DeliveryInfo
        {
            public SourceAddressInfo SourceAddress { get; set; }
            public DestinationAddressInfo DestinationAddress { get; set; }
            public ParcelDataInfo ParcelData { get; set; }
            public List<string> CustomerReferences { get; set; }
            public Dictionary<string, object> Services { get; set; }
            public LabelInfo Label { get; set; }
        }

        public class ReturnsInfo
        {
            public DestinationAddressInfo SourceAddress { get; set; }
            public SourceAddressInfo DestinationAddress { get; set; }
            public ParcelDataInfo ParcelData { get; set; }
            public List<string> CustomerReferences { get; set; }
            public LabelInfo Label { get; set; }
        }

        public class SourceAddressInfo
        {
            public string Title { get; set; }
            public string FirstName { get; set; }
            public string LastName { get; set; }
            public string Street { get; set; }
            public string HouseNo { get; set; }
            [JsonProperty("postcode")]
            public string PostCode { get; set; }
            public string City { get; set; }
            public string Country { get; set; }
            public List<string> AddressAdditionalFields { get; set; }
        }

        public class DestinationAddressInfo
        {
            public string Title { get; set; }
            public string FirstName { get; set; }
            public string LastName { get; set; }
            public string Street { get; set; }
            public string HouseNo { get; set; }
            [JsonProperty("postcode")]
            public string PostCode { get; set; }
            public string City { get; set; }
            public string Country { get; set; }
            public string Mobile { get; set; }
            public string Email { get; set; }
            public List<string> AddressAdditionalFields { get; set; }
        }

        public class ParcelDataInfo
        {
            public string PackagingType { get; set; } = "PC";
            public string Weight { get; set; }
        }

        public class LabelInfo
        {
            public string Output { get; set; } = "PDF";
        }
    }
}

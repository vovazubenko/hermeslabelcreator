namespace HermesLabelCreator.Services
{
    public class CreateLabelResponse
    {
        public DeliveryInfo Delivery { get; set; }
        public ReturnsInfo Returns { get; set; }
        public bool Success { get; set; }
        public string Message { get; set; }
        public string ResponseString { get; set; }

        public class DeliveryInfo
        {
            public string LabelBase64 { get; set; }
            public LabelContentInfo LabelContent { get; set; }
        }

        public class ReturnsInfo
        {
            public string LabelBase64 { get; set; }
            public LabelContentInfo LabelContent { get; set; }
        }

        public class LabelContentInfo
        {
            public TradeInfo Import { get; set; }
            public TradeInfo Export { get; set; }
        }

        public class TradeInfo
        {
            public BarcodeInfo[] BarcodeObject { get; set; }
        }

        public class BarcodeInfo
        {
            public string Barcode { get; set; }
            public string BarcodeFormatted { get; set; }
        }
    }
}

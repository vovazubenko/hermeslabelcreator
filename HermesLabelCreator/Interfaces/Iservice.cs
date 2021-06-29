using HermesLabelCreator.Models;
using HermesLabelCreator.Services;

namespace HermesLabelCreator.Interfaces
{
    public interface Iservice
    {
        CreateLabelResponse GenerateShipmentLabel(Shipment shipment);
    }
}

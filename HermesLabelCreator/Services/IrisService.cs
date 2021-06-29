using HermesLabelCreator.Configurations;
using HermesLabelCreator.Interfaces;
using HermesLabelCreator.Models;
using Microsoft.Extensions.Logging;
using Microsoft.Extensions.Options;
using Newtonsoft.Json;
using Newtonsoft.Json.Serialization;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Net;
using System.Net.Http;
using System.Net.Http.Headers;
using System.Text;
using System.Threading.Tasks;

namespace HermesLabelCreator.Services
{
    public class IrisService : Iservice
    {
        private readonly ApplicationConfiguration _configuration;
        private readonly ILogger<IrisService> _logger;

        public IrisService(
            ILogger<IrisService> logger,
            IOptions<ApplicationConfiguration> configuration)
        {
            _logger = logger;
            _configuration = configuration.Value;
        }

        public CreateLabelResponse GenerateShipmentLabel(Shipment shipment)
        {
            CreateLabelResponse result = null;
            string responseString = string.Empty;

            try
            {
                GenerateShipmentLabelRequest request = MapShipmentToRequest(shipment);
                string requestContent = JsonConvert.SerializeObject(request, new JsonSerializerSettings
                {
                    ContractResolver = new DefaultContractResolver
                    {
                        NamingStrategy = new CamelCaseNamingStrategy()
                    },
                    Formatting = Formatting.Indented
                });

                string accessToken = GetAccessToken(shipment.UserID, shipment.Password);
                using (HttpClient client = new HttpClient())
                {
                    client.DefaultRequestHeaders.Authorization =
                        new AuthenticationHeaderValue("Bearer", accessToken);
                    HttpContent httpContent = new StringContent(requestContent, Encoding.UTF8, "application/json");

                    string serviceUrl = _configuration.DispatcherUrl;

                    var response = client.PostAsync(serviceUrl, httpContent).Result;
                    var responseContent = response.Content;
                    responseString = responseContent.ReadAsStringAsync().Result;

                    if (response.IsSuccessStatusCode)
                    {
                        result = JsonConvert.DeserializeObject<CreateLabelResponse>(responseString);
                        result.Success = true;
                        result.ResponseString = responseString;
                    }
                    else
                    {
                        result = new CreateLabelResponse
                        {
                            Success = false,
                            ResponseString = responseString
                        };
                    }
                }
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, ex.Message);
                result = new CreateLabelResponse
                {
                    Success = false,
                    Message = "Exception was thrown during calling IRIS API. Please see system error logs."
                };
            }

            return result;
        }

        private string GetAccessToken(string UserId, string password)
        {
            if (!string.IsNullOrWhiteSpace(_configuration.Login))
            {
                UserId = _configuration.Login;
                password = _configuration.Password;
            }

            string accessToken = string.Empty;
            using (HttpClient client = new HttpClient())
            {
                var authToken = Encoding.ASCII.GetBytes($"{UserId}:{password}");
                client.DefaultRequestHeaders.Authorization = new AuthenticationHeaderValue("Basic",
                        Convert.ToBase64String(authToken));

                string content = string.Format("grant_type=client_credentials&client_id={0}", UserId);
                HttpContent httpContent = new StringContent(content, Encoding.UTF8, "application/x-www-form-urlencoded");

                var response = client.PostAsync(_configuration.AuthUrl, httpContent).Result;

                var responseContent = response.Content;
                string responseString = responseContent.ReadAsStringAsync().Result;
                if (response.IsSuccessStatusCode)
                {
                    AuthorizationResponse authorizationResponse = JsonConvert.DeserializeObject<AuthorizationResponse>(responseString);
                    accessToken = authorizationResponse.AccessToken;
                }
                else
                {
                    throw new Exception(responseString);
                }
            }

            return accessToken;
        }

        private GenerateShipmentLabelRequest MapShipmentToRequest(Shipment shipment)
        {
            if (shipment.ShipperCountry == "DE" && shipment.ShipperZip.Length < 5)
            {
                shipment.ShipperZip = "0" + shipment.ShipperZip;
            }

            if (shipment.ReceiverCountry == "DE" && shipment.ReceiverZip.Length < 5)
            {
                shipment.ReceiverZip = "0" + shipment.ReceiverZip;
            }

            GenerateShipmentLabelRequest request = new GenerateShipmentLabelRequest
            {
                Delivery = new GenerateShipmentLabelRequest.DeliveryInfo
                {
                    SourceAddress = new GenerateShipmentLabelRequest.SourceAddressInfo
                    {
                        Title = shipment.ShipperTitle,
                        FirstName = shipment.ShipperFirstName,
                        LastName = shipment.ShipperLastName,
                        Street = string.IsNullOrWhiteSpace(shipment.ShipperStreetNumber)
                        ? shipment.ShipperStreet
                        : $"{shipment.ShipperStreet} {shipment.ShipperStreetNumber}",
                        HouseNo = shipment.ShipperStreetNumber,
                        City = shipment.ShipperCity,
                        PostCode = shipment.ShipperZip,
                        Country = shipment.ShipperCountry,
                        AddressAdditionalFields = new List<string> { shipment.ShipperAddressAdd }
                    },
                    DestinationAddress = new GenerateShipmentLabelRequest.DestinationAddressInfo
                    {
                        Title = shipment.ReceiverTitle,
                        FirstName = shipment.ReceiverFirstName,
                        LastName = shipment.ReceiverLastName,
                        Street = shipment.ReceiverStreet,
                        HouseNo = shipment.ReceiverStreetNumber,
                        City = shipment.ReceiverCity,
                        PostCode = shipment.ReceiverZip,
                        Country = shipment.ReceiverCountry,
                        Email = shipment.ReceiverEmail,
                        Mobile = shipment.ReceiverMobile,
                        AddressAdditionalFields = new List<string> { shipment.ReceiverAddressAdd }
                    },
                    ParcelData = new GenerateShipmentLabelRequest.ParcelDataInfo
                    {
                        PackagingType = "PC",
                        Weight = shipment.ArticleWeight
                    },
                    CustomerReferences = new List<string>
                    {
                        shipment.OrderNumber,
                        shipment.PuNumber
                    },
                    Services = new Dictionary<string, object>
                    {
                    },
                    Label = new GenerateShipmentLabelRequest.LabelInfo()
                }                
            };

            if (!string.IsNullOrWhiteSpace(shipment.Returns) && shipment.Returns == "1")
            {
                request.Returns = new GenerateShipmentLabelRequest.ReturnsInfo
                {
                    SourceAddress = new GenerateShipmentLabelRequest.DestinationAddressInfo
                    {
                        Title = shipment.ReceiverTitle,
                        FirstName = shipment.ReceiverFirstName,
                        LastName = shipment.ReceiverLastName,
                        Street = shipment.ReceiverStreet,
                        HouseNo = shipment.ReceiverStreetNumber,
                        City = shipment.ReceiverCity,
                        PostCode = shipment.ReceiverZip,
                        Country = shipment.ReceiverCountry,
                        Email = shipment.ReceiverEmail,
                        Mobile = shipment.ReceiverMobile
                    },
                    DestinationAddress = new GenerateShipmentLabelRequest.SourceAddressInfo
                    {
                        Title = shipment.ShipperTitle,
                        FirstName = shipment.ShipperFirstName,
                        LastName = shipment.ShipperLastName,
                        Street = string.IsNullOrWhiteSpace(shipment.ShipperStreetNumber)
                        ? shipment.ShipperStreet
                        : $"{shipment.ShipperStreet} {shipment.ShipperStreetNumber}",
                        HouseNo = shipment.ShipperStreetNumber,
                        City = shipment.ShipperCity,
                        PostCode = shipment.ShipperZip,
                        Country = shipment.ShipperCountry
                    },
                    CustomerReferences = new List<string>
                    {
                        shipment.OrderNumber,
                        shipment.PuNumber
                    },
                    ParcelData = new GenerateShipmentLabelRequest.ParcelDataInfo
                    {
                        PackagingType = "PC",
                        Weight = shipment.ArticleWeight
                    },
                    Label = new GenerateShipmentLabelRequest.LabelInfo()
                };
            }

            if (!string.IsNullOrWhiteSpace(shipment.Kws) && shipment.Kws == "1")
            {
                request.Delivery.ParcelData.PackagingType = "M1";
                request.Delivery.Services.Add("mailDelivery", new object());

                if (request.Returns != null)
                {
                    request.Returns.ParcelData.PackagingType = "M1";
                }
            }
            else
            {
                request.Delivery.Services.Add("homeDelivery", new object());
            }

            if (!string.IsNullOrWhiteSpace(shipment.NextDay) && shipment.NextDay == "1")
            {
                request.Delivery.Services.Add("nextDay", new object());
            }

            if (!string.IsNullOrWhiteSpace(shipment.BulkyGoods) && shipment.BulkyGoods == "1")
            {
                request.Delivery.Services.Add("bulkyGoods", new object());
            }

            return request;
        }
    }
}

using HermesLabelCreator.Models;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Linq;
using System.Reflection;
using System.Text;
using System.Threading.Tasks;

namespace HermesLabelCreator.Helpers
{
    public static class ImportFileParser
    {
        public static Shipment[] ParseShipments(SpreadsheetCellDto[][] rows)
        {
            List<Shipment> shipments = new List<Shipment>();
            var cellColumnNameValueAssociation = GetCellColumnNameValueAssociation(rows[0]);

            for (int i = 1; i < rows.Length; i++)
            {
                Shipment shipment = new Shipment();
                PropertyInfo[] props = typeof(Shipment).GetProperties();
                foreach (var prop in props)
                {
                    DisplayNameAttribute dp = prop.GetCustomAttributes(typeof(DisplayNameAttribute), true).Cast<DisplayNameAttribute>().SingleOrDefault();
                    if (dp != null)
                    {
                        string displayNameValue = dp.DisplayName;
                        if (cellColumnNameValueAssociation.TryGetValue(displayNameValue, out string cellColumnName))
                        {
                            string value = rows[i]
                                .Where(c => c.CellColumnName == cellColumnName)
                                .Select(c => c.Value)
                                .FirstOrDefault();
                            prop.SetValue(shipment, value);
                        }
                    }
                }

                shipments.Add(shipment);
            }


            return shipments.ToArray();
        }
        private static Dictionary<string, string> GetCellColumnNameValueAssociation(SpreadsheetCellDto[] row)
        {
            Dictionary<string, string> cellColumnNameValueAssociation = new Dictionary<string, string>();
            foreach (var cell in row)
            {
                cellColumnNameValueAssociation.Add(cell.Value, cell.CellColumnName);
            }

            return cellColumnNameValueAssociation;
        }
    }
}

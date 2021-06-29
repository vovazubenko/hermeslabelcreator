using Serilog.Events;
using Serilog.Formatting;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace HermesLabelCreator.Formatters
{
    class CsvFormatter : ITextFormatter
    {
        private const string Quote = "\"";

        private const string EscapedQuote = "\"\"";

        private readonly char[] EscapableCharacters = { ',', '"', '\r', '\n' };

        public static readonly Func<string> HeaderFactory = () => "Excel Row Number,Iris Response";

        public void Format(LogEvent logEvent, TextWriter output)
        {
            if (logEvent.Properties.ContainsKey("rowNumber"))
            {
                this.WriteValue(output, logEvent.Properties["rowNumber"].ToString());
            }

            output.Write(",");

            if (logEvent.Properties.ContainsKey("responseMessage"))
            {
                this.WriteValue(output, logEvent.Properties["responseMessage"].ToString());
            }

            output.WriteLine();
        }

        private void WriteValue(TextWriter output, string value)
        {
            bool needsEscaping = value.IndexOfAny(EscapableCharacters) >= 0;

            if (needsEscaping)
            {
                output.Write(Quote);
                output.Write(value.Replace(Quote, EscapedQuote));
                output.Write(Quote);
            }
            else
            {
                output.Write(value);
            }
        }
    }
}

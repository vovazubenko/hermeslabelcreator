namespace Microsoft.Extensions.Logging
{
    public static class SerilogExtensions
    {
        public static void LogInfo(this ILogger logger, int rowNumber, string responseMessage)
        {
            logger.LogInformation("{rowNumber}{responseMessage}", rowNumber, responseMessage);
        }
    }
}

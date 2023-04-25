

namespace GetComInfo.Helpers
{
    /// <summary>
    /// Http Helper class
    /// </summary>
    public static class HttpHelper
    {
        /// <summary>
        /// Get Http response
        /// </summary>
        public static async Task<string> HttpResponse(string line)
        {
            using (var net = new HttpClient())
            {
                var response = await net.GetAsync(line);
                return response.IsSuccessStatusCode ? await response.Content.ReadAsStringAsync() : null;
            }
        }
    }
}

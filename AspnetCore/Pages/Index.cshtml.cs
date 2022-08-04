using Microsoft.AspNetCore.Mvc;
using Microsoft.AspNetCore.Mvc.RazorPages;

using Microsoft.JSInterop;
using System.Net.Http.Headers;

namespace AspnetCore.Pages
{
    public class IndexModel : PageModel
    {
        private readonly ILogger<IndexModel> _logger;

        public IndexModel(ILogger<IndexModel> logger)
        {
            _logger = logger;

            //ClientCredential cc = new ClientCredential("6c8f5ceb-ce5e-4646-8e4a-43bad62d265c", "lR58Q~55Wfht6GQPU8aCf~lcfbenO1p_HkyiQbxk");
            //var context = new AuthenticationContext("https://login.microsoftonline.com/" + "fadaf86c-7a37-40b6-8a15-7ffe204bb3f3");
            //var result = context.AcquireTokenAsync("https://management.azure.com/", cc);
            //if (result == null)
            //{
            //    throw new InvalidOperationException("Failed to obtain the Access token");
            //}
            //string token = result.Result.AccessToken;

           
            // strtoken = token;
            // var userInfoUrl = "https://portal.office.com";
            // var hc = new HttpClient();
            // hc.DefaultRequestHeaders.Authorization = new AuthenticationHeaderValue("Bearer", token);
            // var response = hc.GetAsync(userInfoUrl).Result;
            // dynamic userInfo = response.Content.ReadAsStringAsync().Result;
            //// return userInfo;
        }

        public void OnGet()
        {
        }

        [BindProperty(SupportsGet = true)]
        public string strtoken { get; set; }

        //[JSInvokable]
        //public IActionResult GetTokenString()
        //{

            
        //    return Content(token);
        //}
    }
}
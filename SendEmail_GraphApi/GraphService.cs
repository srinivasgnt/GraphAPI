using Newtonsoft.Json;
using Newtonsoft.Json.Linq;
using System;
using System.Collections.Generic;
using System.Net;
using System.Net.Http;
using System.Net.Http.Headers;
using System.Text;
using System.Threading.Tasks;
using Resources;
using System.IO;
using Microsoft.Graph;

namespace SendEmail_GraphApi
{

    // This sample shows how to:
    //    - Get the current user's email address
    //    - Get the current user's profile photo
    //    - Attach the photo as a file attachment to an email message
    //    - Upload the photo to the user's root drive
    //    - Get a sharing link for the file and add it to the message
    //    - Send the email
    public class GraphService
    {
        // Send an email message from the current user.
        public async Task<string> SendEmail(string accessToken, MessageRequest email)
        {
            string endpoint = "https://graph.microsoft.com/v1.0/me/sendMail";

            using (var client = new HttpClient())
            {
                using (var request = new HttpRequestMessage(HttpMethod.Post, endpoint))
                {
                    request.Headers.Authorization = new AuthenticationHeaderValue("Bearer", accessToken);

                    request.Content = new StringContent(JsonConvert.SerializeObject(email), Encoding.UTF8, "application/json");
                    using (var response = await client.SendAsync(request))
                    {
                        if (response.IsSuccessStatusCode)
                        {
                            return Resource.Graph_SendMail_Success_Result;
                        }
                        return response.ReasonPhrase;
                    }
                }
            }
        }
    }
}
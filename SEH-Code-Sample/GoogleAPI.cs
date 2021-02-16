using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Net;
using System.IO;
using Newtonsoft.Json;

namespace SEH_Code_Sample
{
    public static class GoogleAPI
    {
        /// <summary>
        /// Interaction logic for MainWindow.xaml
        /// </summary>
        public static List<Items> SearchGoogleImages(List<string> query)
        {
            // Search Engine ID
            string cx = "bf4611498357bc8c9";

            // Google custom search API key
            string apiKey = "AIzaSyBNTVBbjGHJUVkNz9MIT7qa739aibJ0As8";

            List<Items> items = new List<Items>();

            // %20 is a space in the api call
            // Insert %20 in between all words in the query
            string apiQuery = String.Join("%20", query.ToArray());

            // Create Google API call
            string googleUrl = "https://customsearch.googleapis.com/customsearch/v1?key=";
            googleUrl += apiKey;
            googleUrl += "&cx=" + cx;
            googleUrl += "&q=" + apiQuery;
            googleUrl += "&searchType=image";
            googleUrl += "&num=10";
            googleUrl += "&imgSize=MEDIUM";

            try
            {               
                // Make web request 
                var request = WebRequest.Create(googleUrl);
                HttpWebResponse response = (HttpWebResponse)request.GetResponse();
                Stream dataStream = response.GetResponseStream();
                StreamReader reader = new StreamReader(dataStream);
                string responseString = reader.ReadToEnd();
                dynamic jsonData = JsonConvert.DeserializeObject(responseString);

                // Decipher json data
                foreach (var item in jsonData.items)
                {
                    Items temp = new Items();
                    temp.title = item["title"];
                    temp.link = item["link"];
                    temp.snippet = item["snippet"];
                    temp.thumbnailLink = item["image"]["thumbnailLink"];
                    temp.thumbnailHeight = item["image"]["thumbnailHeight"];
                    temp.thumbnailWidth = item["image"]["thumbnailWidth"];

                    items.Add(temp);
                }
            }
            catch (Exception exception)
            {
                return null;
            }
            return items;
        }
    }
}

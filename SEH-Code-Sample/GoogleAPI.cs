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

            // %20 is a space in the api call
            // Insert %20 in between all words in the query
            string apiQuery = String.Join("%20", query.ToArray());

            // Create Google API call
            string googleUrl = "https://customsearch.googleapis.com/customsearch/v1?key=";
            googleUrl += apiKey;
            googleUrl += "&cx=" + cx;
            googleUrl += "&q=" + apiQuery;
            googleUrl += "&searchType=image";
            googleUrl += "&num=9";     // Max number that can be searched is 10
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

                return jsonToItems(jsonData);
            }
            catch (Exception exception)
            {
                return null;
            }
        }

        public static List<Items> jsonToItems(dynamic jsonData)
        {
            List<Items> items = new List<Items>();

            // Check if there are any Items 
            if (jsonData.searchInformation.totalResults != "0")
            {
                // Decompose json data
                foreach (var item in jsonData.items)
                {
                    items.Add(new Items
                    {
                        title = item.title,
                        link = item.link,
                        snippet = item.snippet,
                        thumbnailLink = item.image.thumbnailLink,
                        thumbnailHeight = item.image.thumbnailHeight,
                        thumbnailWidth = item.image.thumbnailWidth
                    });
                }
            }

            return items;
        }
    }
}

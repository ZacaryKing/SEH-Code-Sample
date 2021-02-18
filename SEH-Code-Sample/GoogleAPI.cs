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
        /// Makes web request to Google API using query words given
        /// </summary>
        public static List<Items> SearchGoogleImages(List<string> query)
        {
            // Open up config to get google API parameters
            StreamReader reader = new StreamReader(Path.Combine(Environment.CurrentDirectory, "Resources", "config.json"));
            dynamic jsonData = JsonConvert.DeserializeObject(reader.ReadToEnd());
            reader.Close();

            string googleUrl = jsonToGoogleUrl(jsonData, query);

            try
            {
                // Make web request 
                var request = WebRequest.Create(googleUrl);
                HttpWebResponse response = (HttpWebResponse)request.GetResponse();
                Stream dataStream = response.GetResponseStream();
                reader = new StreamReader(dataStream);
                string responseString = reader.ReadToEnd();
                jsonData = JsonConvert.DeserializeObject(responseString);
                reader.Close();

                return jsonToItems(jsonData);
            }
            catch (Exception exception)
            {
                return null;
            }
        }

        /// <summary>
        /// Grabs Google API parameters from config.json file and creates Google API Url
        /// </summary>
        public static string jsonToGoogleUrl(dynamic jsonData, List<string> query)
        {
            // Insert %20 (space) in between all words in the query
            string apiQuery = String.Join("%20", query.ToArray());
            string googleUrl = "https://customsearch.googleapis.com/customsearch/v1?key=";
            
            googleUrl += jsonData.apiKey;
            googleUrl += "&cx=" + jsonData.searchEngineID;
            googleUrl += "&q=" + apiQuery;
            googleUrl += "&searchType=" + jsonData.searchType;
            googleUrl += "&num=" + jsonData.num;     // Max number that can be searched is 10
            googleUrl += "&imgSize=" + jsonData.imgSize;

            return googleUrl;
        }

        /// <summary>
        /// Converts results from Google API call to a list of Items
        /// </summary>
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

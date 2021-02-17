using System;
using System.Text;
using System.Collections.Generic;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using System.IO;
using Newtonsoft.Json;
using SEH_Code_Sample;

namespace SEH_Code_Sample_Tests
{
    /// <summary>
    /// Summary description for GoogleAPITest
    /// </summary>
    [TestClass]
    public class GoogleAPITest
    {
        string resourcesPath;

        public GoogleAPITest()
        {
            resourcesPath = Path.Combine(Environment.CurrentDirectory, "..", "..", "Resources");
        }
        
        // UnitOfWork_StateUnderTest_ExpectedBehavior
        // Arrange, Act, Assert
        [TestMethod]
        public void JsonToItems_JsonWith1Item_Item()
        {
            // Arrange
            StreamReader reader = new StreamReader(Path.Combine(resourcesPath, "oneSearchResult.json"));
            
            dynamic jsonData = JsonConvert.DeserializeObject(reader.ReadToEnd());

            // Act
            List<Items> items = GoogleAPI.jsonToItems(jsonData);

            // Assert
            Assert.AreEqual("Banana bread tops the list of most searched for recipes in ...", items[0].title);
            Assert.AreEqual("https://i.dailymail.co.uk/1s/2020/05/08/16/28165848-0-image-a-51_1588951482497.jpg", items[0].link);
            Assert.AreEqual("Banana bread tops the list of most searched for recipes in ...", items[0].snippet);
            Assert.AreEqual("https://encrypted-tbn0.gstatic.com/images?q=tbn:ANd9GcSNlrx0wzJ7YfD5cVIIQvOnrEirC-lfVesmhILv4T1D2WyNTHonIHAX&s", items[0].thumbnailLink);
            Assert.AreEqual(70, items[0].thumbnailHeight);
            Assert.AreEqual(117, items[0].thumbnailWidth);
        }

        [TestMethod]
        public void JsonToItems_JsonWith5Items_5Items()
        {
            // Arrange
            StreamReader reader = new StreamReader(Path.Combine(resourcesPath, "fiveSearchResults.json"));
            dynamic jsonData = JsonConvert.DeserializeObject(reader.ReadToEnd());

            // Act
            List<Items> items = GoogleAPI.jsonToItems(jsonData);

            // Assert
            Assert.AreEqual("Google Slides: Free Online Presentations for Personal Use", items[0].title);
            Assert.AreEqual("https://blogs.shu.ac.uk/shutel/files/2014/08/GSlides.png", items[0].link);
            Assert.AreEqual("Google Slides: Free Online Presentations for Personal Use", items[0].snippet);
            Assert.AreEqual("https://encrypted-tbn0.gstatic.com/images?q=tbn:ANd9GcTJD2kAqjGymXR2mNOk4StjJQavh1nwHdkRrFCOuLYrhgkJVz-8JRhbcA&s", items[0].thumbnailLink);
            Assert.AreEqual(116, items[0].thumbnailHeight);
            Assert.AreEqual(116, items[0].thumbnailWidth);

            Assert.AreEqual("Google", items[1].title);
            Assert.AreEqual("https://www.google.com/images/branding/googlelogo/2x/googlelogo_color_160x56dp.png", items[1].link);
            Assert.AreEqual("googlelogo_color_160x56dp.png", items[1].snippet);
            Assert.AreEqual("https://encrypted-tbn0.gstatic.com/images?q=tbn:ANd9GcQu7ggDqQyXuxqO0_rzP-_dRCIY-Ul34jHadQSgkoV-z2gcxrj5rYO0Wg&s", items[1].thumbnailLink);
            Assert.AreEqual(41, items[1].thumbnailHeight);
            Assert.AreEqual(118, items[1].thumbnailWidth);

            Assert.AreEqual("Google Input Tools", items[2].title);
            Assert.AreEqual("https://www.google.com/inputtools/images/hero-video.jpg", items[2].link);
            Assert.AreEqual("Google Input Tools", items[2].snippet);
            Assert.AreEqual("https://encrypted-tbn0.gstatic.com/images?q=tbn:ANd9GcSt9y6Ig6Uxvz5Y_zTtBIkVNnvihzlrABhf-GQiwBsfHbSkQKpgHhN3&s", items[2].thumbnailLink);
            Assert.AreEqual(75, items[2].thumbnailHeight);
            Assert.AreEqual(122, items[2].thumbnailWidth);

            Assert.AreEqual("30th Anniversary of PAC-MAN", items[3].title);
            Assert.AreEqual("https://www.google.com/logos/2002/dilberttwo.gif", items[3].link);
            Assert.AreEqual("30th Anniversary of PAC-MAN", items[3].snippet);
            Assert.AreEqual("https://encrypted-tbn0.gstatic.com/images?q=tbn:ANd9GcSnhk3QXrOnD9DHoRYrgiKbK1zOZxfA3UgG-e1EAge5cGPsXH4kxAqCDVc&s", items[3].thumbnailLink);
            Assert.AreEqual(31, items[3].thumbnailHeight);
            Assert.AreEqual(133, items[3].thumbnailWidth);

            Assert.AreEqual("Google Slides: Free Online Presentations for Personal Use", items[4].title);
            Assert.AreEqual("https://lh3.ggpht.com/9rwhkrvgiLhXVBeKtScn1jlenYk-4k3Wyqt1PsbUr9jhGew0Gt1w9xbwO4oePPd5yOM=w300", items[4].link);
            Assert.AreEqual("Google Slides: Free Online Presentations for Personal Use", items[4].snippet);
            Assert.AreEqual("https://encrypted-tbn0.gstatic.com/images?q=tbn:ANd9GcRf6eeLn_JMfOKbyBwUtVCJLBqnL08o7jw-nx4ocjUmeZOwqa0N5q9WlA&s", items[4].thumbnailLink);
            Assert.AreEqual(116, items[4].thumbnailHeight);
            Assert.AreEqual(116, items[4].thumbnailWidth);
        }

        [TestMethod]
        public void JsonToItems_JsonWithNoItems_EmptyItemsList()
        {
            // Arrange
            StreamReader reader = new StreamReader(Path.Combine(resourcesPath, "noSearchResults.json"));
            dynamic jsonData = JsonConvert.DeserializeObject(reader.ReadToEnd());

            // Act
            List<Items> items = GoogleAPI.jsonToItems(jsonData);

            // Assert
            Assert.AreEqual(0, items.Count);
        }
    }
}

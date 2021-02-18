using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Navigation;
using System.IO;

namespace SEH_Code_Sample
{

    /// <summary>
    /// Interaction logic for MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window
    {

        public MainWindow()
        {
            InitializeComponent();
        }

        #region Properties

        private List<string> query = new List<string>();
        private List<ImageToUse> imagesToPPT = new List<ImageToUse>();

        #endregion

        #region Events

        private void button_Click(object sender, RoutedEventArgs e)
        {
            toggleButton((Button)sender);
        }
        
        private void exitButton_Click(object sender, RoutedEventArgs e)
        {
            this.Close();
        }

        private void helpButton_Click(object sender, RoutedEventArgs e)
        {
            helpInformation();
        }

        private void clearButton_Click(object sender, RoutedEventArgs e)
        {
            clearAppData();
        }
      
        private void pullImagesButton_Click(object sender, RoutedEventArgs e)
        {
            if (getQueryWords())
            {
                string status = poplulateImageGrid();

                if (status != "Success")
                    MessageBox.Show(status, "Error:");
            }
        }

        private void generateSlideButton_Click(object sender, RoutedEventArgs e)
        {
            string status = PPModifications.generatePowerPointSlide(titleTextArea.Text, stringFromRichTextBox(bodyTextArea), imagesToPPT);

            if (status != "Success")
                MessageBox.Show(status);
        }


        #endregion

        #region Private Methods

        /// <summary>
        /// Grabs index number of button found in its name and toggles corresponding button along with
        /// the usage of the image displayed on it.
        /// </summary>
        private void toggleButton(Button selectedButton)
        {
            int imageIndex = 0;
            Int32.TryParse(selectedButton.Name.Substring(1), out imageIndex);

            if (imagesToPPT[imageIndex].use)
            {
                imagesToPPT[imageIndex].use = false;
                selectedButton.BorderThickness = new Thickness(0, 0, 0, 0);
            }
            else
            {
                imagesToPPT[imageIndex].use = true;
                selectedButton.BorderThickness = new Thickness(10, 10, 10, 10);
            }
        }

        /// <summary>
        /// Display message box with all the functionality of the project.
        /// </summary>
        private void helpInformation()
        {
            string helpInfo = "";

            helpInfo += "This Project allows you to enter a title and body. Words from the title and bolded words inside the body will be put inside a query ";
            helpInfo += "once pressing the \"Pull Images\" button. The queried words will then be inserted in a Google custom search API call and populate ";
            helpInfo += "the grid below the body with relevant images. You may then press the \"Generate Slide\" button to insert the Title, Body, and selected ";
            helpInfo += "images to a new PowerPoint slide which will be appended to a PowerPoint called \"Sample.pptx\". \"Sample.pptx\" is located in the ";
            helpInfo += "directory the program is running.";

            MessageBox.Show(helpInfo, "Usage:");
        }

        /// <summary>
        /// Clear the contents in Title, Body, and imageGrid.
        /// </summary>
        private void clearAppData()
        {
            query.Clear();
            imagesToPPT.Clear();
            titleTextArea.Clear();
            bodyTextArea.Document.Blocks.Clear();
            imageGrid.Children.Clear();
        }

        /// <summary>
        /// Grabs words in titleTextArea and bolded words in bodyTextArea and puts them in query.
        /// </summary>
        private bool getQueryWords()
        {
            query.Clear();

            // Get all words from titleTextArea
            string[] temp = titleTextArea.Text.Split();

            // Clear strings with only whitespace, then insert in query
            temp = temp.Where(x => !string.IsNullOrEmpty(x)).ToArray();
            query = temp.ToList();

            // Put bolded words from body into the query
            extractBoldedWordsFromBody(bodyTextArea);

            if (query.Count == 0)
            {
                MessageBox.Show("There are no words in the title text box or any words bolded in the body text box.", "Error:");
                return false;
            }
            return true;
        }

        /// <summary>
        /// Converts RichTextBox text to string made up of xaml data. Searches for bolded words and inserts them into query.
        /// </summary>
        private void extractBoldedWordsFromBody(RichTextBox rtb)
        {
            // Convert RichTextbox text to string made of xaml data
            System.Windows.Documents.TextRange tr =
               new System.Windows.Documents.TextRange(rtb.Document.ContentStart, rtb.Document.ContentEnd);
            MemoryStream ms = new MemoryStream();
            tr.Save(ms, DataFormats.Xaml);
            string xamlText = ASCIIEncoding.Default.GetString(ms.ToArray());

            const string startDelimeter = "<Run FontWeight=\"Bold\">";
            const string endDelimeter = "</Run>";

            int start = 0;
            int end = 0;
            string[] temp;

            // Find all bolded sections in xamlText
            while ((start = xamlText.IndexOf(startDelimeter, end)) != -1)
            {
                // Find start and end position to each bolded sections
                start += startDelimeter.Length;
                end = xamlText.IndexOf(endDelimeter, start);

                // Separate words
                temp = xamlText.Substring(start, end - start).Split();

                // Remove any whitespace entries
                temp = temp.Where(x => !string.IsNullOrEmpty(x)).ToArray();

                // Join new words with query
                query.AddRange(temp);

                end += endDelimeter.Length;
            }
        }

        /// <summary>
        /// Populates imageGrid with buttons with images as backgrounds. Images were taken from Google custom search API call
        /// given words from the query.
        /// </summary>
        private string poplulateImageGrid()
        {
            imageGrid.Children.Clear();
            imagesToPPT.Clear();

            List<Items> items = GoogleAPI.SearchGoogleImages(query);

            if (items == null)
                return "Something went wrong while connecting to Google API.";
            else if (!items.Any())
                return "No search results returned.";

            // Max number of buttons to make
            int max;
            if (items.Count < 9)
                max = items.Count;
            else
                max = 9;

            // Traverse through each row
            for (int rowIndex = 0, i = 0; (rowIndex < imageGrid.RowDefinitions.Count) && (i < max); rowIndex++)
            {
                // Traverse through each column
                for (int columnIndex = 0; (columnIndex < imageGrid.ColumnDefinitions.Count) && (i < max); columnIndex++, i++)
                {
                    // Store images
                    imagesToPPT.Add(new ImageToUse {
                        bitmapImage = new BitmapImage(new Uri(items[i].thumbnailLink)),
                        use = false
                    });

                    Button button = new Button();

                    // Set button properties and background with image found from Google
                    setGridButtonProperties(button, i, imagesToPPT[i].bitmapImage);

                    // Populate each cell with button
                    Grid.SetRow(button, rowIndex);
                    Grid.SetColumn(button, columnIndex);
                    imageGrid.Children.Add(button);
                }
            }
            return "Success";
        }

        /// <summary>
        /// Returns text inside RichTextBox
        /// </summary>
        private string stringFromRichTextBox(RichTextBox rtb)
        {
            System.Windows.Documents.TextRange textRange = new System.Windows.Documents.TextRange(rtb.Document.ContentStart, rtb.Document.ContentEnd);
            return textRange.Text;
        }

        /// <summary>
        /// Programattically sets imageGrid button's properties
        /// </summary>
        private void setGridButtonProperties(Button button, int buttonNumber, BitmapImage bitmapImage)
        {
            // Set properties
            button.Name = "n" + buttonNumber.ToString();
            button.HorizontalAlignment = HorizontalAlignment.Center;
            button.VerticalAlignment = VerticalAlignment.Center;
            button.BorderBrush = Brushes.BlueViolet;
            button.BorderThickness = new Thickness(0, 0, 0, 0);

            // Set Background Properties
            var brush = new ImageBrush();
            brush.ImageSource = bitmapImage;
            brush.Stretch = Stretch.Uniform;
            button.Background = brush;

            // Size of Grid Cell
            button.Height = imageGrid.ActualHeight / imageGrid.RowDefinitions.Count;
            button.Width = imageGrid.ActualWidth / imageGrid.ColumnDefinitions.Count;

            // Set Click Event
            button.Click += new RoutedEventHandler(button_Click);
        }

        #endregion

    }
}

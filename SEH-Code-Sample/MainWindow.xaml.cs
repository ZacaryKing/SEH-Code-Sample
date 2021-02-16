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
using System.Windows.Shapes;
using Microsoft.Office.Interop.PowerPoint;
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
            generatePowerPointSlide();
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
            helpInfo += "images to a PowerPoint slide to a PowerPoint presentation called \"Sample.pptx\" in the directory the program is running.";

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

            if ((items == null) || (!items.Any()))
                return "Something went wrong trying to connect to Google API.";

            // Traverse through each row
            for (int rowIndex = 0, i = 0; (rowIndex < imageGrid.RowDefinitions.Count) && (i < items.Count); rowIndex++)
            {
                // Traverse through each column
                for (int columnIndex = 0; (columnIndex < imageGrid.ColumnDefinitions.Count) && (i < items.Count); columnIndex++, i++)
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
            button.Height = 128;
            button.Width = 230;

            // Set Click Event
            button.Click += new RoutedEventHandler(button_Click);
        }

        /// <summary>
        /// Generates PowerPoint Slide with given Title, Body and selected images
        /// </summary>
        private void generatePowerPointSlide()
        {
            Microsoft.Office.Interop.PowerPoint.Application pptApplication = new Microsoft.Office.Interop.PowerPoint.Application();
            Microsoft.Office.Interop.PowerPoint.Presentation pptPresentation = pptApplication.Presentations.Add(Microsoft.Office.Core.MsoTriState.msoTrue);

            Microsoft.Office.Interop.PowerPoint.CustomLayout customLayout =
                pptPresentation.SlideMaster.CustomLayouts[Microsoft.Office.Interop.PowerPoint.PpSlideLayout.ppLayoutText];

            // Create new Slide
            Microsoft.Office.Interop.PowerPoint.Slides slides = pptPresentation.Slides;

            // Append new slide to end of presentation
            Microsoft.Office.Interop.PowerPoint._Slide slide = slides.AddSlide(1, customLayout);
            
            Microsoft.Office.Interop.PowerPoint.TextRange objText;

            // Add title from titleTextArea
            objText = slide.Shapes[1].TextFrame.TextRange;
            objText.Text = titleTextArea.Text;
            objText.Font.Name = "Arial";
            objText.Font.Size = 32;

            // Add body from bodyTextArea
            objText = slide.Shapes[2].TextFrame.TextRange;
            objText.Text = stringFromRichTextBox(bodyTextArea);
            objText.Font.Name = "Arial";
            objText.Font.Size = 20;
        
            Microsoft.Office.Interop.PowerPoint.Shapes shapes = slide.Shapes;
            string imagePath = System.IO.Path.Combine(Environment.CurrentDirectory, "image.jpg");
            int imageX = 50;

            // Go through all images, if image was highlighted in grid, post it inside PPT
            foreach (ImageToUse imageToUse in imagesToPPT)
            {
                if (imageToUse.use)
                {
                    // Save all images being used
                    var encoder = new PngBitmapEncoder();
                    encoder.Frames.Add(BitmapFrame.Create((BitmapSource)imageToUse.bitmapImage));
                    using (FileStream stream = new FileStream(imagePath, FileMode.Create)) encoder.Save(stream);

                    // Insert all images into PPT
                    shapes.AddPicture(imagePath, Microsoft.Office.Core.MsoTriState.msoFalse, Microsoft.Office.Core.MsoTriState.msoTrue, imageX, 350, 100, 100);
                    imageX += 100;

                    // Delete all images when done with them
                    File.Delete(imagePath);
                }
            }

            try
            {
                // Save Power Point
                string powerPointPath = System.IO.Path.Combine(Environment.CurrentDirectory, "ppSample.pptx");
                pptPresentation.SaveAs(powerPointPath, Microsoft.Office.Interop.PowerPoint.PpSaveAsFileType.ppSaveAsDefault, Microsoft.Office.Core.MsoTriState.msoTrue);
            }
            catch (Exception exception)
            {
                MessageBox.Show("Problem while trying to save PowerPoint:\n" + exception.ToString(), "Error:");
            }

            // pptPresentation.Close();
            // pptApplication.Quit();
        }

        #endregion

    }
}

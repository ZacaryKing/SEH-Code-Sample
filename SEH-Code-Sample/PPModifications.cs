using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Microsoft.Office.Interop.PowerPoint;
using System.IO;
using System.Windows.Media.Imaging;

namespace SEH_Code_Sample
{
    public static class PPModifications
    {
        /// <summary>
        /// Generates PowerPoint Slide with given Title, Body and selected images
        /// </summary>
        public static string generatePowerPointSlide(string titleText, string bodyText, List<ImageToUse> imagesToPPT)
        {
            string status = "Success";
            string powerPointPath = System.IO.Path.Combine(Environment.CurrentDirectory, "ppSample.pptx");
            Microsoft.Office.Interop.PowerPoint.Application pptApplication = new Microsoft.Office.Interop.PowerPoint.Application();
            Microsoft.Office.Interop.PowerPoint.Presentation pptPresentation;

            if (File.Exists(powerPointPath))
            {
                // Append to existing PowerPoint
                Microsoft.Office.Interop.PowerPoint.Presentations pres = pptApplication.Presentations;
                pptPresentation = pres.Open(powerPointPath);
            }
            else
            {
                // Create new PowerPoint
                pptPresentation = pptApplication.Presentations.Add(Microsoft.Office.Core.MsoTriState.msoTrue);
            }

            Microsoft.Office.Interop.PowerPoint.CustomLayout customLayout =
                    pptPresentation.SlideMaster.CustomLayouts[Microsoft.Office.Interop.PowerPoint.PpSlideLayout.ppLayoutText];

            // Acquire slides from presentation
            Microsoft.Office.Interop.PowerPoint.Slides slides = pptPresentation.Slides;

            // Append new slide to end of presentation
            Microsoft.Office.Interop.PowerPoint._Slide slide = slides.AddSlide(slides.Count + 1, customLayout);
            
            Microsoft.Office.Interop.PowerPoint.TextRange objText;

            // Add title from titleTextArea
            objText = slide.Shapes[1].TextFrame.TextRange;
            objText.Text = titleText;
            objText.Font.Name = "Arial";
            objText.Font.Size = 32;

            // Add body from bodyTextArea
            objText = slide.Shapes[2].TextFrame.TextRange;
            objText.Text = bodyText;
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
                pptPresentation.SaveAs(powerPointPath, Microsoft.Office.Interop.PowerPoint.PpSaveAsFileType.ppSaveAsDefault, Microsoft.Office.Core.MsoTriState.msoTrue);
            }
            catch (Exception exception)
            {
                status = "Problem while trying to save PowerPoint:\n" + exception.ToString();
            }

            return status;
            // pptPresentation.Close();
            // pptApplication.Quit();
        }
    }
}

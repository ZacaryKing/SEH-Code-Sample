## SEH-Code-Sample

#### Environment Used:

Microsoft Visual Studio Community 2015 Version 14.0.25431.01 Update 3

#### Usage: 

This Project allows you to enter a title and body. Words from the title and bolded words inside the body will be put inside a query once pressing the "Pull Images" button. The queried words will then be inserted in a Google custom search API call and populate the grid below the body with relevant images. You may then press the "Generate Slide" button to insert the Title, Body, and selected images to a PowerPoint slide to a PowerPoint presentation called "Sample.pptx" in the directory the program is running.

#### Notes:

I left the Google search engine ID and Google API Key in the project. These along with the other Google API parameters can be altered in the SearchGoogleImages method inside the GoogleAPI class.
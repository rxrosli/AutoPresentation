
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Reflection;
using Office = Microsoft.Office.Core;
using PowerPoint = Microsoft.Office.Interop.PowerPoint;



namespace AutoPresentation
{
    static class Program
    {
        static void Main(string[] args)
        {
            AutoPresentation();
            // Clean up the unmanaged PowerPoint COM resources by forcing a  
            // garbage collection as soon as the calling function is off the  
            // stack (at which point these objects are no longer rooted). 
            GC.Collect();
            GC.WaitForPendingFinalizers();
            // GC needs to be called twice in order to get the Finalizers called  
            // - the first time in, it simply makes a list of what is to be  
            // finalized, the second time in, it actually is finalizing. Only  
            // then will the object do its automatic ReleaseComObject. 
            GC.Collect();
            GC.WaitForPendingFinalizers();
        }

        private static void AutoPresentation()
        {
            try
            {
                // Create an instance of Microsoft PowerPoint and make it  
                // invisible. 

                PowerPoint.Application oPowerPoint = new PowerPoint.Application();


                // By default PowerPoint is invisible, till you make it visible: 
                // oPowerPoint.Visible = Office.MsoTriState.msoFalse; 


                // Create a new Presentation. 

                PowerPoint.Presentation oPre = oPowerPoint.Presentations.Add(Office.MsoTriState.msoFalse);
                Console.WriteLine("Presentation created");

                //Data Parse
                
                string BaseDir = Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location);

                Console.WriteLine(BaseDir + "\\data.csv");
                List<Presentee> collection = new List<Presentee>();
                using (StreamReader sr = new StreamReader(BaseDir + "\\data.csv"))
                {
                    string line;
                    while ((line = sr.ReadLine()) != null)
                    {
                        //Console.WriteLine(line);
                        string[] values = line.Split(',');
                        collection.Add(new Presentee { 
                        LastName = Convert.ToString(values[0]),
                        FirstName = Convert.ToString(values[1]),
                        Initial = Convert.ToString(values[2]),
                        Faculty = Convert.ToString(values[3]),
                        Directory = Convert.ToString(values[4])
                        });
                    }
                }
              
                // Insert a new Slide and add some text to it.
                PowerPoint.Slide oSlide;
                PowerPoint.TextFrame oFrame;
                PowerPoint.TextRange oText;
                int SlideCount = 0;
                foreach (Presentee set in collection)
                {
                    oSlide = oPre.Slides.Add(++SlideCount, PowerPoint.PpSlideLayout.ppLayoutBlank);

                    // Last Name
                    oSlide.Shapes.AddTextbox(
                        Office.MsoTextOrientation.msoTextOrientationHorizontal,
                        430,
                        150,
                        500,
                        50
                    );

                    //First Name & Initials
                    oSlide.Shapes.AddTextbox(
                        Office.MsoTextOrientation.msoTextOrientationHorizontal,
                        430,
                        150 + 115,
                        500,
                        50
                    );

                    //Faculty
                    oSlide.Shapes.AddTextbox(
                        Office.MsoTextOrientation.msoTextOrientationHorizontal,
                        430,
                        150 + 115 + 120,
                        500,
                        50
                    );

                    // Alignment
                    oFrame = oSlide.Shapes[1].TextFrame;
                    oFrame.VerticalAnchor = Office.MsoVerticalAnchor.msoAnchorMiddle;
                    oFrame = oSlide.Shapes[2].TextFrame;
                    oFrame.VerticalAnchor = Office.MsoVerticalAnchor.msoAnchorMiddle;
                    oFrame = oSlide.Shapes[3].TextFrame;
                    oFrame.VerticalAnchor = Office.MsoVerticalAnchor.msoAnchorMiddle;

                    // Text Format & Content
                    oText = oSlide.Shapes[1].TextFrame.TextRange;
                    oText.ParagraphFormat.Alignment = PowerPoint.PpParagraphAlignment.ppAlignCenter;
                    oText.Text = set.LastName + ",";
                    oText.Font.Name = "Palatino Linotype";
                    oText.Font.Bold = Office.MsoTriState.msoTrue;
                    oText.Font.Size = 72;

                    oText = oSlide.Shapes[2].TextFrame.TextRange;
                    oText.ParagraphFormat.Alignment = PowerPoint.PpParagraphAlignment.ppAlignCenter;
                    oText.Text = set.FirstName + " " + set.Initial + ".";
                    oText.Font.Name = "Palatino Linotype";
                    oText.Font.Size = 64;

                    oText = oSlide.Shapes[3].TextFrame.TextRange;
                    oText.ParagraphFormat.Alignment = PowerPoint.PpParagraphAlignment.ppAlignCenter;
                    oText.Text = set.Faculty;
                    oText.Font.Name = "Palatino Linotype";
                    oText.Font.Color.RGB = 1689855;                    
                    oText.Font.Size = 40;


                    //Picture
                    oSlide.Shapes.AddPicture(
                        set.Directory,
                        Office.MsoTriState.msoFalse,
                        Office.MsoTriState.msoTrue,
                        30,
                        30,
                        380,
                        500
                    );

                    Console.WriteLine("Slide {0}: {1},{2} {3}. ",SlideCount, set.LastName,set.FirstName,set.Initial);
                }

                // Save the presentation as a pptx file and close it. 

                Console.WriteLine("Save and close the presentation");

                string fileName = BaseDir + "\\output.pptx";
                oPre.SaveAs(fileName,
                    PowerPoint.PpSaveAsFileType.ppSaveAsOpenXMLPresentation,
                    Office.MsoTriState.msoTriStateMixed);
                oPre.Close();

                // Quit the PowerPoint application. 

                Console.WriteLine("Quit the PowerPoint application");
                oPowerPoint.Quit();
                Console.ReadKey();
            }
            catch (Exception ex)
            {
                Console.WriteLine("Program.AutomatePowerPoint throws the error: {0}", ex.Message);
                Console.ReadKey();
            }
        }
    }

    class Presentee
    {   
        public string LastName { get; set; }
        public string FirstName { get; set; }
        public string Initial { get; set; }
        public string Faculty { get; set; }
        public string Directory { get; set; }
    }
}
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using DocumentFormat.OpenXml.Office.Drawing;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Presentation;
using Shape = DocumentFormat.OpenXml.Presentation.Shape;

namespace PPT_To_Latex
{
    class Program
    {
        static void Main(string[] args)
        {
            // http://msdn.microsoft.com/en-us/library/bb448854.aspx

            bool includeHidden = true;
            string pptxfilename = "test2.pptx";
            string latexfilename = "test.tex";

            ConvertPptx2Tex(pptxfilename, includeHidden, latexfilename);
        }

        private static void ConvertPptx2Tex(string pptxfilename, bool includeHidden, string latexfilename)
        {
            using (PresentationDocument presentationDocument = PresentationDocument.Open(pptxfilename, false))
            {
                var presentationPart = presentationDocument.PresentationPart;
                var presentation = presentationPart.Presentation;

                //Count slides
                Console.WriteLine("Slides counts={0}", SlidesCount(includeHidden, presentationPart));
                int imageCount = 0;
                var fileresult = File.CreateText(latexfilename);

                foreach (SlideId slideId in presentation.SlideIdList)
                {
                   


                    String relId = slideId.RelationshipId.Value;

                    SlidePart slide = (SlidePart)presentation.PresentationPart.GetPartById(relId);
                   
                    if (!includeHidden) //check if we got hidden slide, of so, skip
                    {
                        if (slide.Slide.Show!=null &&  slide.Slide.Show.HasValue && slide.Slide.Show.Value == false)
                            continue;
                    }



                    if (slide.SlideLayoutPart.SlideLayout.Type == SlideLayoutValues.SectionHeader)
                    {
                        Debug.WriteLine("%%%%%%%%%%NEWSECTION%%%%%%%%%%%%");
                        fileresult.WriteLine(@"\section{" + GetSlideTitle(slide) + "}");
                        fileresult.WriteLine();
                    }


                    Debug.WriteLine("\n\n\n********************************");
                    //Get title
                    var paragraphTexttit = GetSlideTitle(slide);
                    Debug.WriteLine("\t\t" + paragraphTexttit.ToString());
                    Debug.WriteLine("----------------------");

                    fileresult.WriteLine();
                    fileresult.WriteLine(@"\begin{frame}");
                    fileresult.WriteLine(@"\frametitle{" + paragraphTexttit + "}");
                    fileresult.WriteLine();

                    int previndent = 0;
                    bool firstitemdone = false;
                    foreach (var paragraph in slide.Slide.Descendants<DocumentFormat.OpenXml.Drawing.Paragraph>().Skip(1))
                    {
                        //http://msdn.microsoft.com/en-us/library/ee922775(v=office.14).aspx
                        int currentIndentLevel = 0;

                        if (paragraph.ParagraphProperties != null)
                        {
                            if (paragraph.ParagraphProperties.HasAttributes)
                            {
                                try
                                {
                                    string lvl = paragraph.ParagraphProperties.GetAttribute("lvl", "").Value;
                                    currentIndentLevel = int.Parse(lvl);
                                }
                                catch
                                {
                                    //Ignore
                                }
                            }
                        }

                        StringBuilder paragraphText = new StringBuilder();
                        // Iterate through the lines of the paragraph.
                        foreach (var text in paragraph.Descendants<DocumentFormat.OpenXml.Drawing.Text>())
                        {
                            paragraphText.Append(text.Text);
                        }

                        if (paragraphText.Length > 0)
                        {
                            if (firstitemdone == false)
                            {
                                WriteWithIndent(fileresult, @"\begin{itemize}[<+->]", currentIndentLevel);
                                firstitemdone = true;
                            }
                            if (previndent > currentIndentLevel)
                            {
                                WriteWithIndent(fileresult, @"\end{itemize}", currentIndentLevel + 1);
                            }
                            else if (previndent < currentIndentLevel)
                            {
                                WriteWithIndent(fileresult, @"\begin{itemize}", currentIndentLevel);
                            }
                            if (!paragraphText.ToString().Contains(@"artesis 20"))
                                //TODO: this could be added as external input (e.g. filter file/blacklist)
                                WriteWithIndent(fileresult, @"\item " + paragraphText, currentIndentLevel);
                            Debug.WriteLine(paragraphText.ToString());
                        }
                        previndent = currentIndentLevel;
                    }

                    //Get all the images!!! 
                    foreach (var pic in slide.Slide.Descendants<Picture>())
                    {
                        try
                        {
                            string extension;

                            //Extract correct image part and extenion
                            var imagePart = ExtractImage(pic, slide, out extension);

                            //Write text item
                            imageCount++;
                            string imagefilename = string.Format("{0}{1}.{2}", "image", imageCount, extension);
                            fileresult.WriteLine(@"\begin{figure}[h] \begin{center}");
                            fileresult.WriteLine("\t" + @"\includegraphics[width=0.5\textwidth]{" + imagefilename + "}");
                            fileresult.WriteLine(@"\end{center} \end{figure}");

                            //Save image to file
                            WriteImageToFile(imagePart, imagefilename);
                        }
                        catch (Exception ex)
                        {
                            Console.WriteLine("Error with an image");
                        }
                    }

                    if (firstitemdone == true)
                    {
                        fileresult.WriteLine(@"\end{itemize}");
                    }

                    fileresult.WriteLine(@"\end{frame}");
                }
                fileresult.Close();
            }
        }

        private static void WriteImageToFile(ImagePart imagePart, string filename)
        {
            System.Drawing.Image img = System.Drawing.Image.FromStream(imagePart.GetStream());
            img.Save(filename);
        }

        private static ImagePart ExtractImage(Picture pic, SlidePart slide, out string extension)
        {
            // First, get relationship id of image
            string rId = pic.BlipFill.Blip.Embed.Value;

            ImagePart imagePart = (ImagePart)slide.GetPartById(rId);

            // Get the original file name.
            Debug.WriteLine("$$Image:" + imagePart.Uri.OriginalString);
            extension = "bmp";
            if (imagePart.ContentType.Contains("jpeg") || imagePart.ContentType.Contains("jpg"))
                extension = "jpg";
            else if (imagePart.ContentType.Contains("png"))
                extension = "png";
            return imagePart;
        }

        private static void WriteWithIndent(StreamWriter fileresult, string stringtowrite, int indentlevel)
        {
            if (indentlevel < 0)
                indentlevel = 0;
            StringBuilder sb = new StringBuilder();
            for (int i = 0; i < indentlevel; i++)
            {
                sb.Append("\t");
            }
            sb.Append(stringtowrite);

            fileresult.WriteLine(sb.ToString());
        }

        private static StringBuilder GetSlideTitle(SlidePart slide)
        {
            var shapes = from shape in slide.Slide.Descendants<Shape>()
                         where IsTitleShape(shape)
                         select shape;
            StringBuilder paragraphTexttit = new StringBuilder();
            string paragraphSeparator = null;
            foreach (var shape in shapes)
            {
                // Get the text in each paragraph in this shape.
                foreach (var paragraph in shape.TextBody.Descendants<DocumentFormat.OpenXml.Drawing.Paragraph>())
                {
                    // Add a line break.
                    paragraphTexttit.Append(paragraphSeparator);

                    foreach (var text in paragraph.Descendants<DocumentFormat.OpenXml.Drawing.Text>())
                    {
                        paragraphTexttit.Append(text.Text);
                    }

                    paragraphSeparator = "\n";
                }
            }
            return paragraphTexttit;
        }

        private static int SlidesCount(bool includeHidden, PresentationPart presentationPart)
        {
            int slidesCount = 0;
            if (includeHidden)
            {
                slidesCount = presentationPart.SlideParts.Count();
            }
            else
            {
                var slides =
                    presentationPart.SlideParts.Where(
                        (s) => (s.Slide != null) && ((s.Slide.Show == null) || (s.Slide.Show.HasValue && s.Slide.Show.Value)));
                slidesCount = slides.Count();
            }
            return slidesCount;
        }

        // Determines whether the shape is a title shape.
        private static bool IsTitleShape(Shape shape)
        {
            var placeholderShape = shape.NonVisualShapeProperties.ApplicationNonVisualDrawingProperties.GetFirstChild<PlaceholderShape>();
            if (placeholderShape != null && placeholderShape.Type != null && placeholderShape.Type.HasValue)
            {
                switch ((PlaceholderValues)placeholderShape.Type)
                {
                    // Any title shape.
                    case PlaceholderValues.Title:

                    // A centered title.
                    case PlaceholderValues.CenteredTitle:
                        return true;

                    default:
                        return false;
                }
            }
            return false;
        }
    }
}

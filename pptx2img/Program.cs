using System;
using System.IO;
using CommandLine;
using CommandLine.Text;
using Microsoft.Office.Core;
using Microsoft.Office.Interop.PowerPoint;

namespace pptx2img
{
    class Program
    {
        public class Options
        {
            [Option('f', "FileName", Required = true, HelpText = "PowerPoint (*.pptx) file name")]
            public string FileName { get; set; }

            [Option('o', "OutDir", Required = true, HelpText = "Output directory")]
            public string OutDir { get; set; }

            [HelpOption]
            public string GetUsage()
            {
                return HelpText.AutoBuild(this, _ => HelpText.DefaultParsingErrorsHandler(this, _));
            }
        }

        static void Main(string[] args)
        {
            Options options = new Options();
            if (!Parser.Default.ParseArguments(args, options))
            {
                return;
            }

            Application application = null;
            try
            {
                application = new Application();
                var fileName = Path.GetFullPath(options.FileName);
                var presentation = application.Presentations.Open(fileName,
                    MsoTriState.msoFalse, MsoTriState.msoFalse, MsoTriState.msoFalse);

                if (!Directory.Exists(options.OutDir))
                {
                    Directory.CreateDirectory(options.OutDir);
                }

                var fileNameWithoutExtension = Path.GetFileNameWithoutExtension(options.FileName);
                int i = 1;
                foreach (Slide slide in presentation.Slides)
                {
                    var shape = slide.Shapes.Range(slide.Shapes.GetIndices()).Group();
                    shape.Export(Path.Combine(Path.GetFullPath(options.OutDir), $"{fileNameWithoutExtension}_{i}.png"), PpShapeFormat.ppShapeFormatPNG);
                    i++;
                }
            }
            catch (Exception e)
            {
                Console.WriteLine(e);
                application?.Quit();
            }
        }
    }
}

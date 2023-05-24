using Docentric.Documents.ObjectModel;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace DocentricConsoleApplicationNetCore
{

    public abstract class ReportBase
    {
        public string Name { get; protected set; }
        public string Title { get; protected set; }
        public string Description { get; protected set; }
        public string HelpTopicPage { get; protected set; }
        public string DirectoryPath { get; private set; }

        // Constructor
        public ReportBase()
        {
            // Default output format is docx.
            SaveOptions = SaveOptions.Word;

            // Set the directory path of the executing assembly
            string directoryPath = System.IO.Path.GetDirectoryName(System.Reflection.Assembly.GetExecutingAssembly().Location);
        }
        /// <summary>
        /// Generates the report for the current template.
        /// </summary>
        /// <returns>File full name of the genarated report document.</returns>
        public abstract string GenerateReport();


        /// <summary>
        /// Opens the report template file for this example.
        /// </summary>
        /// <returns>Returns the file stream of the report template file.</returns>
        protected Stream GetReportTemplate()
        {
            if (!File.Exists(TemplateFileName))
            {
                throw new Exception(string.Format("Report template '{0}' is not found", TemplateFileName));
            }
            return File.Open(TemplateFileName, FileMode.Open, FileAccess.Read, FileShare.ReadWrite);
        }


        // ApplicationPath
        protected string ApplicationPath
        {
            get
            {
                var url = System.IO.Path.GetDirectoryName(System.Reflection.Assembly.GetExecutingAssembly().GetName().CodeBase);
                var mainpath = url.Replace("\\bin\\Debug", "").Replace("file:\\", "").Replace("net7.0", "");
                return mainpath;
            
            }
        }


        // TemplateFileName
        public string TemplateFileName
        {
            get
            {
                return System.IO.Path.Combine(ApplicationPath, "Resources\\" + Name + "\\Template.docx");
            }
        }

        // SaveOptions
        public SaveOptions SaveOptions
        {
            get;
            set;
        }


        protected string OutputDocumentFileExtension
        {
            get
            {
                if (SaveOptions is WordSaveOptions) return ".docx";
                else if (SaveOptions is PdfSaveOptions) return ".pdf";
                else if (SaveOptions is XpsSaveOptions) return ".xps";
                else throw new Exception("Unsupported 'SaveOptions' object.");
            }
        }
    }

}

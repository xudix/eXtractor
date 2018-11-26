using System;
using System.IO;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Xml;
using System.Xml.Serialization;

namespace eXtractor
{
    [Serializable]
    public class ExtractionRequestModel
    {

        public DateTime StartDateTime { get; set; }


        public DateTime EndDateTime { get; set; }


        public string FilePath { get; set; }

        [XmlArray(ElementName ="Tags")]
        public string[] SelectedTags { get; set; }

        [XmlArray(ElementName = "DataFiles")]
        public string[] SelectedFiles { get; set; }

        public int Interval { get; set; }

        public int Resolution { get; set; }

        private static string settingFile = "eXtractorSettings.xml";

        public ExtractionRequestModel()
        {
            StartDateTime = DateTime.Now;
            EndDateTime = DateTime.Now;
            FilePath = String.Empty;
            Resolution = 1000;
            Interval = 1;
        }

        public static ExtractionRequestModel CreateExtractionRequestModel()
        {
            try
            {
                using(StreamReader sr = new StreamReader(new FileStream(settingFile, FileMode.Open, FileAccess.Read, FileShare.ReadWrite)))
                {
                    return new XmlSerializer(typeof(ExtractionRequestModel)).Deserialize(sr) as ExtractionRequestModel;
                }
            }
            catch
            {
                return new ExtractionRequestModel();
            }
        }

        public void SaveSettings()
        {
            using (StreamWriter sw = new StreamWriter(settingFile))
            {
                XmlSerializer xmlSerializer = new XmlSerializer(typeof(ExtractionRequestModel));
                xmlSerializer.Serialize(sw, this);
            }
        }

        /// <summary>
        /// Extract data as specified and export the data to a file
        /// This method will create an ExtractedData object and call its WriteToFile method
        /// </summary>
        public void Export()
        {
            if (SelectedTags != null && SelectedFiles != null)
            {
                (new ExtractedData(this)).WriteToFile(StartDateTime, EndDateTime, "csv", FilePath);
                SaveSettings();
            }
        }

    }
}

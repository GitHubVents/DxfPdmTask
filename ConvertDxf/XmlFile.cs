using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using ExportToXMLLib;
using Patterns.Observer;

namespace ConvertDxf
{
    public class XmlFile
    {
        public void DownloadXml(string path)
        {
            try
            {
                //MessageBox.Show("Статус: Выгрузка XML файлов");
                Export export = new Export(path);
                export.XML();
                MessageObserver.Instance.SetMessage("XML no EX.");
            }
            catch (Exception ex)
            {
                MessageObserver.Instance.SetMessage("XML EX: " + ex.StackTrace);
                MessageObserver.Instance.SetMessage("Failed to upload XML with exception: " + Environment.NewLine + ex.Message);
            }
        }
    }
}
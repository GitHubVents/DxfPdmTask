using System;
using System.Windows.Forms;
using Patterns.Observer;
using SolidWorksLibrary;

namespace ConvertDxf.Log
{
    static class Program
    {
        static bool toSend = false;

        public static void Instance_ReceivedMessage(MessageEventArgs massage)
        {
            try
            {
                string pathToLog = System.Environment.GetFolderPath(System.Environment.SpecialFolder.ApplicationData) + @"\ExportPartToDXF\Log.txt";

                Logger.Instance.ToLog($"Time:{massage.time} Message: {massage.Message}");

                if (massage.Type == MessageType.Error)
                {
                    toSend = true;
                    MessageObserver.Instance.SetMessage(massage.Message);
                }
                //if (massage.Message == "End upload.\n" && toSend)
                //{
                //    //SolidWorksAdapter.OutLookSendMeALog(pathToLog, "Some exception from ExportToDXF");
                //}



            }
            catch (Exception ex)
            {
                MessageObserver.Instance.SetMessage(ex.ToString());
            }
        }
    }
}
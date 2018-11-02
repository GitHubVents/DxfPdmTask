using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using ConvertDxf.Log;
using EPDM.Interop.epdm;
using Patterns.Observer;

namespace ConvertDxf
{
    [Guid("2F65E2BD-F835-4892-84B0-7348469C06FA"), ComVisible(true)]
    public class Class1 : IEdmAddIn5
    {
        private string tempAppFolder { get { return Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData) + "\\ExportPartToDXF"; } }

        public void GetAddInInfo(ref EdmAddInInfo poInfo, IEdmVault5 poVault, IEdmCmdMgr5 poCmdMgr)
        {
            poInfo.mbsAddInName = "DownloadDxf";
            poInfo.mbsCompany = "SOLIDWORKS Corporation";
            poInfo.mbsDescription = "Adds menu command items";
            poInfo.mlAddInVersion = 1;

            poInfo.mlRequiredVersionMajor = 10;
            poInfo.mlRequiredVersionMinor = 0;

            poCmdMgr.AddCmd(100, "Добавить Dxf-file", (int)EdmMenuFlags.EdmMenu_OnlyMultipleSelection);
            poCmdMgr.AddCmd(100, "Добавить Dxf-file", (int)EdmMenuFlags.EdmMenu_OnlySingleSelection);
            poCmdMgr.AddHook(EdmCmdType.EdmCmd_TaskRun);
            poCmdMgr.AddHook(EdmCmdType.EdmCmd_TaskSetup);

        }

        public void OnCmd(ref EdmCmd poCmd, ref EdmCmdData[] ppoData)
        {
            switch (poCmd.meCmdType)
            {
                case EdmCmdType.EdmCmd_TaskRun:
                    OnTaskRun(ref poCmd, ref ppoData);
                    break;
                case EdmCmdType.EdmCmd_TaskSetup:
                    OnTaskSetup(ref poCmd, ref ppoData);
                    break;
                case EdmCmdType.EdmCmd_Menu:
                   OnMenu(ref poCmd, ref ppoData);
                    break;
            }
        }

        private void OnMenu(ref EdmCmd poCmd, ref EdmCmdData[] ppoData)
        {
            try
            {
                Logger.Instance.RootDirectory = tempAppFolder;
                MessageObserver.Instance.ReceivedMessage += Program.Instance_ReceivedMessage;

                for (int i = 0; i < ppoData.Length; i++)
                {
                    if (((EdmCmdData)ppoData.GetValue(0)).mlObjectID1 != 0)
                    {
                        IEdmVault5 vault = new EdmVault5();

                        vault = (IEdmVault5)poCmd.mpoVault;
                        IEdmObject5 folderObject = vault.GetObject(EdmObjectType.EdmObject_Folder,
                            ((EdmCmdData)ppoData.GetValue(i)).mlObjectID3);
                        IEdmFolder5 ef = (IEdmFolder5)folderObject;
                        IEdmObject5 fileObject = vault.GetObject(EdmObjectType.EdmObject_File,
                            ((EdmCmdData)ppoData.GetValue(i)).mlObjectID1);
                        string fullpath = ef.LocalPath + "\\" + fileObject.Name;

                        DxfLoad dxfl = new DxfLoad();
                        var sp = dxfl.GetSpecification(fullpath, "");
                        dxfl.UpLoadDxf(sp);
                        
                        //XmlFile xf = new XmlFile();
                        //xf.DownloadXml(fullpath);
                    }

                    SolidWorksPdmAdapter.Instance.KillProcsses("SLDWORKS");
                    //MessageObserver.Instance.SetMessage("End upload.\n");
                }
            }
            catch (Exception ex)
            {
                MessageObserver.Instance.SetMessage(ex.Message);
                throw;
            }
        }

        private void OnTaskRun(ref EdmCmd poCmd, ref EdmCmdData[] ppoData)
        {
            IEdmTaskInstance inst = default(IEdmTaskInstance);
            inst = (IEdmTaskInstance)poCmd.mpoExtra;

            IEdmVault5 vault = new EdmVault5();

            try
            {
                for (int i = 0; i < ppoData.Length; i++)
                {
                    if (((EdmCmdData)ppoData.GetValue(0)).mlObjectID1 != 0)
                    {
                        vault = (IEdmVault5)poCmd.mpoVault;
                        IEdmObject5 folderObject = vault.GetObject(EdmObjectType.EdmObject_Folder, ((EdmCmdData)ppoData.GetValue(i)).mlObjectID2);
                        IEdmFolder5 ef = (IEdmFolder5)folderObject;
                        IEdmObject5 fileObject = vault.GetObject(EdmObjectType.EdmObject_File,
                            ((EdmCmdData)ppoData.GetValue(i)).mlObjectID1);
                        string fullpath = ef.LocalPath + "\\" + fileObject.Name;

                        DxfLoad dxfl = new DxfLoad();
                        var sp = dxfl.GetSpecification(fullpath, "");
                        dxfl.UpLoadDxf(sp);
                        XmlFile xf = new XmlFile();
                        xf.DownloadXml(fullpath);
                    }

                    SolidWorksPdmAdapter.Instance.KillProcsses("SLDWORKS");
                    MessageObserver.Instance.SetMessage("End upload.\n");
                }
            }
            catch (System.Runtime.InteropServices.COMException ex)
            {
                inst.SetStatus(EdmTaskStatus.EdmTaskStat_DoneFailed, ex.ErrorCode, "The test task failed!");
            }
            catch (Exception ex)
            {
                MessageObserver.Instance.SetMessage(ex.Message);
            }
        }

        private void OnTaskSetup(ref EdmCmd poCmd, ref EdmCmdData[] ppoData)
        {
            try
            {
                
                IEdmTaskProperties props = default(IEdmTaskProperties);
                props = (IEdmTaskProperties)poCmd.mpoExtra;
              
            }
            catch (System.Runtime.InteropServices.COMException ex)
            {
                MessageObserver.Instance.SetMessage("HRESULT = 0x" + ex.ErrorCode.ToString("X") + ex.Message);
            }
            catch (Exception ex)
            {
                MessageObserver.Instance.SetMessage(ex.Message);
            }
        }
    }
}
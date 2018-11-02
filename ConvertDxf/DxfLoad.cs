using System;
using System.Collections.Generic;
using System.Deployment.Application;
using System.IO;
using System.Linq;
using System.Runtime.Remoting.Contexts;
using System.Windows.Forms;
using ConvertDxf.Log;
using ConvertDxf.Models;
using ConvertDxf.Models.ORM;
using ExportToXMLLib;
using Patterns.Observer;
using SolidWorksLibrary.Builders.Dxf;

namespace ConvertDxf
{
    public class DxfLoad
    {
        internal static Properties.Settings settings { get { /*Properties.Settings.Default.Reload( ); */return Properties.Settings.Default; } }
        private string tempAppFolder { get { return Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData) + "\\ExportPartToDXF"; } }

        IEnumerable<Specification> specificationsQuery;

        public DxfLoad()
        {
           
            if (ApplicationDeployment.IsNetworkDeployed)
            {
                Version version = ApplicationDeployment.CurrentDeployment.CurrentVersion;
                MessageObserver.Instance.SetMessage($"v. {version.ToString()}");

            }

            try
            {
                SolidWorksPdmAdapter.Instance.BoomId = settings.BoomId;
            }
            catch (Exception ex)
            {
                MessageObserver.Instance.SetMessage(ex.ToString());
            }
            PdmLogin();
        }

        public void PdmLogin()
        {
            try
            {
                if (!String.IsNullOrEmpty(settings.VaultName))

                {
                    SolidWorksPdmAdapter.Instance.AuthoLogin(settings.VaultName);
                }
            }
            catch (Exception ex)
            {
                MessageObserver.Instance.SetMessage("Не верное имя PDM вида: " + settings.VaultName + "\n" + ex.Message);
            }
        }

        public Specification[] GetSpecification(string filePath, string configuration)
        {
           try
            {
                var parts = AdapterPdmDB.Instance.Parts;
                var bomShell = SolidWorksPdmAdapter.Instance.GetBomShell(filePath, configuration);
                if (bomShell != null)
                {
                    specificationsQuery = (from eachBom in bomShell
                        join eachPart in parts on new { id = eachBom.IdPdm, ver = eachBom.Version, conf = eachBom.Configuration }
                        equals new { id = (int)eachPart.IDPDM, ver = eachPart.Version, conf = eachPart.ConfigurationName }
                        into specS

                        from spec in specS.DefaultIfEmpty()
                        select new Specification
                        {
                            Description = eachBom.Description,
                            PartNumber = eachBom.PartNumber,
                            Version = eachBom.Version,
                            Configuration = eachBom.Configuration,
                            Idpdm = eachBom.IdPdm,
                            Partition = eachBom.Partition,

                            Bend = (spec == null ? 0 : spec.Bend),
                            PaintX = (spec == null ? 0 : spec.PaintX),
                            PaintY = (spec == null ? 0 : spec.PaintY),
                            PaintZ = (spec == null ? 0 : spec.PaintZ),
                            WorkpieceX = (spec == null ? 0 : spec.WorkpieceX),
                            WorkpieceY = (spec == null ? 0 : spec.WorkpieceY),
                            SurfaceArea = (spec == null ? 0 : spec.SurfaceArea),
                            IsDxf = (spec == null || spec.DXF != "1") ? false : true,
                            FileName = eachBom.FileName,
                            FilePath = Path.Combine(eachBom.FolderPath, eachBom.FileName),
                            Thickness = (spec == null ? 0 : spec.Thickness)
                        });


                    return specificationsQuery.ToList().Where(each => each.FileName.ToLower().Contains(".sldprt")
                                                                      && (each.Partition == string.Empty || each.Partition == "Детали")).ToArray();
                }
            }
            catch (Exception ex)
            {
                MessageObserver.Instance.SetMessage(ex.ToString());

            }
            return null;
        }

        public void UpLoadDxf(Specification[] specifications)
        {
            DxfBulder dxfBulder = DxfBulder.Instance;
            string tempDxfFolder = tempAppFolder;
            int countIterations = 0;

            tempDxfFolder = Path.Combine(tempDxfFolder, "DXF");

            dxfBulder.DxfFolder = tempDxfFolder;
            dxfBulder.FinishedBuilding += DxfBulder_FinishedBuilding;
            //MessageBox.Show("Статус: Получение файлов");
            SolidWorksPdmAdapter.Instance.DownLoadFile(specifications);

            //IEnumerable<FileModelPdm> fileModelsOfSpecification = Specification.ConvertToFileModels(specifications);

            //MessageBox.Show("Статус: Выгрузка DXF файлов");

            if (specifications != null)
            {
                specifications = specifications.GroupBy(each => each.FilePath).Select(each => new Specification
                {
                    Description = each.First().Description,
                    PartNumber = each.First().PartNumber,
                    Version = each.First().Version,
                    Configuration = each.First().Configuration,
                    Idpdm = each.First().Idpdm,
                    Bend = each.First().Bend,
                    PaintX = each.First().PaintX,
                    PaintY = each.First().PaintY,
                    WorkpieceX = each.First().WorkpieceX,
                    PaintZ = each.First().PaintZ,
                    WorkpieceY = each.First().WorkpieceY,
                    SurfaceArea = each.First().SurfaceArea,
                    IsDxf = each.First().IsDxf,
                    FileName = each.First().FileName,
                    FilePath = each.First().FilePath,
                    Thickness = each.First().Thickness,
                }).ToArray();

                foreach (var eachSpec in specifications)
                {
                    if (!eachSpec.IsDxf && Path.GetExtension(eachSpec.FileName).ToUpper() == ".SLDPRT")
                    {
                        //MessageBox.Show("Статус: Выгрузка DXF файла " + eachSpec.FileName + "-" + eachSpec.Configuration);
                        dxfBulder.Build(eachSpec.FilePath, eachSpec.Idpdm, eachSpec.Version, eachSpec.Configuration);

                        try
                        {
                            //MessageBox.Show("Статус: Выгрузка XML файлов");
                            Export export = new Export(eachSpec.FilePath);
                            export.XML();
                            MessageObserver.Instance.SetMessage("XML no EX.");
                        }
                        catch (Exception ex)
                        {
                            MessageObserver.Instance.SetMessage("XML EX: " + ex.StackTrace);
                            MessageObserver.Instance.SetMessage("Failed to upload XML with exception: " + Environment.NewLine + ex.Message);
                        }

                        countIterations++;
                        MessageObserver.Instance.SetMessage($"Created {countIterations} new dxf files to temp folder");
                    }
               }
            }
    
            countIterations = 0;

            #region  load dxf as binary from database  and save as dxf file
            foreach (var item in specificationsQuery)
            {
                try
                {
                    if (AdapterPdmDB.Instance.IsDxf(item.Idpdm, item.Configuration, item.Version))
                    {
                        byte[] binary = AdapterPdmDB.Instance.GetDXF(item.Idpdm, item.Configuration, item.Version);
                        string fileName = Path.GetFileNameWithoutExtension(item.FileName).Replace("ВНС-", "");
                        
                        fileName = DXF.DxfNameBuild(fileName, item.Configuration) + "-" + item.Thickness.ToString().Replace(",", ".");

                        if (Path.GetExtension(item.FileName.ToUpper()) == ".SLDPRT")
                        {
                            BinaryToDxfFile(binary, fileName, Path.Combine(tempDxfFolder));

                        }
                        else
                        {
                            BinaryToDxfFile(binary, fileName, Path.Combine(tempDxfFolder, Path.GetFileNameWithoutExtension(item.FileName)));
                        }
                        countIterations++;
                        
                    }
                }
                catch (Exception ex)
                {
                    MessageObserver.Instance.SetMessage($"Failed to save new dxf files to destination folder.\t" + item.FilePath + Environment.NewLine + ex.Message + Environment.NewLine + ex.StackTrace);
                    throw ex;
                }
             }

            #endregion

            #region clear temp dxf directory
            var files = Directory.GetFiles(tempDxfFolder);
            foreach (var file in files)
            {
                File.Delete(file);
            }
            #endregion

            //MessageBox.Show("Статус: Выгрузка DXF файлов завершена. Количество выгруженых файлов " + countIterations);
            //countIterations = 0;

        }
        private void BinaryToDxfFile(byte[] inputBinary, string fileName, string directory)
        {
            string path = DXF.DxfNameBuild(fileName, "");

            try
            {
                path = Path.Combine(directory, $"{fileName}.dxf");
            }
            catch (Exception ex)
            {
                MessageObserver.Instance.SetMessage("BinaryToDxfFile failed to create path.\nfileName " + fileName + ";    directory: " + directory);
                throw ex;
            }

            if (!Directory.Exists(directory))
                Directory.CreateDirectory(directory);
            File.WriteAllBytes(path, inputBinary);
        }
        private void DxfBulder_FinishedBuilding(DataToExport dataToExport)
        {
            try
            {
                AdapterPdmDB.Instance.UpDateCutList(dataToExport.Configuration,
                    dataToExport.DXFByte,
                    dataToExport.WorkpieceX,
                    dataToExport.WorkpieceY,
                    dataToExport.Bend,
                    dataToExport.Thickness,
                    dataToExport.Version,
                    dataToExport.PaintX,
                    dataToExport.PaintY,
                    dataToExport.PaintZ,
                    dataToExport.IdPdm,
                    dataToExport.SurfaceArea,
                    dataToExport.MaterialID
                );

            }
            catch (Exception ex)
            {
                MessageObserver.Instance.SetMessage(ex.ToString());
            }
        }
    }
}
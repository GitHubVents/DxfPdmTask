using System.Collections.Generic;
using System.ComponentModel;

namespace ConvertDxf.Models
{
    public class Specification : ISpecificationView
    {
        public bool IsDxf { get; set; }
        public string PartNumber { get; set; } // +
        public string Description { get; set; } // 
        public decimal WorkpieceX { get; set; }
        public decimal WorkpieceY { get; set; }
        public int Bend { get; set; }
        public decimal Thickness { get; set; }
        public string Configuration { get; set; }
        public int Version { get; set; }
        public int PaintX { get; set; }
        public int PaintY { get; set; }
        public int PaintZ { get; set; }
        public decimal SurfaceArea { get; set; }
        public string FileName { get; set; }
        public int Idpdm { get; set; }
        public string FilePath { get; set; }
        public string Partition { get; set; }

        public static IEnumerable<ISpecificationView> ConvertToViews(IEnumerable<Specification> specification)
        {
            List<ISpecificationView> resultList = new List<ISpecificationView>();
            foreach (var item in specification)
            {
                resultList.Add(item);
            }
            return resultList;
        }
        public static IEnumerable<FileModelPdm> ConvertToFileModels(IEnumerable<Specification> specification)
        {
            List<FileModelPdm> fileModels = new List<FileModelPdm>();
            foreach (Specification eachSpec in specification)
            {
                fileModels.Add(new FileModelPdm
                {
                    IDPdm = eachSpec.Idpdm,
                    CurrentVersion = eachSpec.Version,
                    FileName = eachSpec.FileName,
                    FolderPath = System.IO.Path.GetDirectoryName(eachSpec.FilePath),
                    Path = eachSpec.FilePath,
                    FolderId = SolidWorksPdmAdapter.Instance.GetFolderId(System.IO.Path.GetDirectoryName(eachSpec.FilePath))
                });
            }
            return fileModels;
        }
    }
    public interface ISpecificationView
    {
        [DisplayName("Текущий DXF статус")]
        bool IsDxf { get; set; }

        [DisplayName("Обозначение")]
        string PartNumber { get; set; }
        [DisplayName("Наименование")]
        string Description { get; set; }
        [DisplayName("Конфигурация")]
        string Configuration { get; set; }
        [DisplayName("Версия")]
        int Version { get; set; }
    }
}
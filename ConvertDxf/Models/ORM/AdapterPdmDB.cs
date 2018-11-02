﻿using System;
using System.Collections.Generic;
using System.Data.Linq;
using ConvertDxf.Properties;
using Patterns;
using Patterns.Observer;

namespace ConvertDxf.Models.ORM
{
    class AdapterPdmDB : Singeton<AdapterPdmDB>
    {
        SWPlusDataContext DataContext { get { return new SWPlusDataContext(Settings.Default.DBConnectionString); } }
        protected AdapterPdmDB() : base()
        {

        }
        public IEnumerable<View_Part> Parts
        {
            get
            {
                return DataContext.View_Parts;
            }
        }
        public void UpDateCutList(

            string configuration,
            byte[] DXFByte,
            decimal workpieceX,
            decimal workpieceY,
            int bend,
            decimal thickness,
            int version,
            int paintX,
            int paintY,
            int paintZ,
            int IdPdm,
            decimal surfaceArea,
            int? materialID = null
        )

        {
            try
            {
                DataContext.DXFUpDateCutList(
                    workpieceX,
                    workpieceY,
                    bend,
                    thickness,
                    configuration,
                    version,
                    paintX,
                    paintY,
                    paintZ,
                    IdPdm,
                    materialID,
                    surfaceArea,
                    new Binary(DXFByte));
                ////////////////////////////////////////////////////////////////////////////////////////
                MessageObserver.Instance.SetMessage("Successfully added the cut list to database.   " + IdPdm.ToString(), MessageType.System);
            }
            catch (Exception exception)
            {
                MessageObserver.Instance.SetMessage("Failed added the cut list to database; Exception message: " + exception.ToString(), MessageType.Error);
            }
        }
        public bool IsDxf(int IdPDM, string configuration, int version)
        {
            return DataContext.DXFCheck(IdPDM, configuration, version) != 0 ? true : false;
        }

        public byte[] GetDXF(int IdPDM, string configuration, int version)
        {
            Binary binaryDxf = null;
            DataContext.DXF_GET(IdPDM, configuration, version, ref binaryDxf);
            return binaryDxf.ToArray();
        }
    }
}
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.IO;
using System.IO.Compression;
using System.Windows.Forms;

namespace Charge_Slip.Services
{
    class ZipServices
    {
        public void ZipFileS(string _processby, Main main)
        {

            string sPath = Application.StartupPath + "\\" + Main.outputFolder;
            string dPath = Application.StartupPath + "\\AFT_" + main.batchfile + "_" + _processby + ".zip";
            //DeleteMDB(main.FileName + ".mdb");
            //DeleteMDB("ChargeSlip.mdb");
            DeleteFiles(".zip",Application.StartupPath);
            //DeleteMDB(".mdb");
            ZipFile.CreateFromDirectory(sPath, dPath);
          // Ionic.Zip.ZipFile zips = new Ionic.Zip.ZipFile(dPath);
            //Adding order file to zip file
            //zips.AddItem(Application.StartupPath + "\\MDB");
            //zips.Save();

        }
        public void DeleteFiles(string _ext,string _path)
        {

            DirectoryInfo di = new DirectoryInfo(_path);
            FileInfo[] files = di.GetFiles("*" + _ext)
                     .Where(p => p.Extension == _ext).ToArray();
            foreach (FileInfo file in files)
            {
                file.Attributes = FileAttributes.Normal;
                File.Delete(file.FullName);
            }
        }

        public void DeleteMDB(string _ext)
        {
            DirectoryInfo di = new DirectoryInfo(Application.StartupPath + "\\MDB");
            FileInfo[] files = di.GetFiles("*" + _ext)
                     .Where(p => p.Extension == _ext).ToArray();
            foreach (FileInfo file in files)
            {
                file.Attributes = FileAttributes.Normal;
                File.Delete(file.FullName);
            }
        }
        public static void CopyMDB(string _processby, Main main, string _filename)
        {
            string dPath = Application.StartupPath +"\\" + Main.outputFolder + "\\" + _filename +".mdb";
            string sPath = Application.StartupPath + "\\MDB\\" + _filename  +".mdb";
            File.Copy(sPath, dPath, true);
            

        }
        public void CopyZipFile(string _processby, Main main, string _IpAddress)
        {
            if (main.printerFileOutputFolder == "UNION")
            {
                string dPath = "\\\\" + _IpAddress + "\\captive\\Zips\\" + main.printerFileOutputFolder + "\\VOUCHER\\AFT_" + main.batchfile + "_" + _processby + ".zip";
                string sPath = Application.StartupPath + "\\AFT_" + main.batchfile + "_" + _processby + ".zip";
                File.Copy(sPath, dPath, true);

            }
            else
            { 
            string dPath = "\\\\" + _IpAddress + "\\captive\\Zips\\" + main.printerFileOutputFolder + "\\CHARGE SLIP\\AFT_" + main.batchfile + "_" + _processby + ".zip";
            string sPath = Application.StartupPath + "\\AFT_" + main.batchfile + "_" + _processby + ".zip";
            File.Copy(sPath, dPath, true);
            }
        }
        public void CopyZipFileM(string _processby, Main main, string _IpAddress)
        {
            string dPath = "\\\\" + _IpAddress + "\\captive\\Zips\\Metro\\CHARGE SLIP\\AFT_" + main.batchfile + "_" + _processby + ".zip";
            string sPath = Application.StartupPath + "\\AFT_" + main.batchfile + "_" + _processby + ".zip";
            File.Copy(sPath, dPath, true);

        }
        public void CopyMDB(string _processby, Main main)
        {

            //string dPath = "\\\\192.168.0.254\\PrinterFiles\\" + main.printerFileOutputFolder+ "\\TEST\\"+DateTime.Now.Year+"\\CHARGE SLIP\\"+main.FileName +main.checktype+ ".mdb";
            //string sPath = Application.StartupPath + "\\"+Main.outputFolder+"\\" + main.FileName + main.checktype + ".mdb";
            //File.Copy(sPath, dPath, true);
            if (main.printerFileOutputFolder == "METRO")
            {
                string dPath = "\\\\192.168.0.254\\PrinterFiles\\" + main.printerFileOutputFolder + "\\CHARGE SLIP\\" + main.FileName  + ".mdb";
                string sPath = Application.StartupPath + "\\" + Main.outputFolder + "\\" + main.FileName  + ".mdb";
                File.Copy(sPath, dPath, true);
            }
            else if (main.printerFileOutputFolder == "UNION")
            {
                string dPath = "\\\\192.168.0.254\\PrinterFiles\\" + main.printerFileOutputFolder + "\\VOUCHER\\" + main.FileName  + ".mdb";
                string sPath = Application.StartupPath + "\\" + Main.outputFolder + "\\" + main.FileName  + ".mdb";
                File.Copy(sPath, dPath, true);
            }
            else
            {
                string dPath = "\\\\192.168.0.254\\PrinterFiles\\" + main.printerFileOutputFolder + "\\CHARGE SLIP\\" + main.FileName + ".mdb";
                string sPath = Application.StartupPath + "\\" + Main.outputFolder + "\\" + main.FileName + ".mdb";
                File.Copy(sPath, dPath, true);
            }

        }
        public void CopyPacking(string _processby, Main main)
        {
            //string dPath = "\\\\192.168.0.254\\Packing\\" + Main.outputFolder + "\\TEST\\";
            //string sPath = Application.StartupPath + "\\" + Main.outputFolder + "\\Packing.dbf";
            //{
            //    Directory.CreateDirectory(dPath + main.batchfile);
            //}
            //string dpath2 = dPath + "\\" + main.batchfile;
            string dPath = "";
            //File.Copy(sPath, dpath2 + "\\Packing.dbf", true);
            string sPath = Application.StartupPath + "\\" + Main.outputFolder + "\\Packing.dbf";
            if (main.packingOutputFolder == "UNION")
                dPath = "\\\\192.168.0.254\\Packing\\" + main.packingOutputFolder + "\\VOUCHER\\";
            else
            dPath = "\\\\192.168.0.254\\Packing\\" + main.packingOutputFolder + "\\CHARGE SLIP\\";

            
            {
                Directory.CreateDirectory(dPath + main.batchfile);
            }
            string dpath2 = dPath + "\\" + main.batchfile;

            File.Copy(sPath, dpath2 + "\\Packing.dbf", true);

        }
    }
}

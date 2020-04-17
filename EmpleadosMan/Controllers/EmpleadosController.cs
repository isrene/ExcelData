using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.Mvc;
using Excel = Microsoft.Office.Interop.Excel;
using System.IO;
using EmpleadosMan.Models;
using System.Globalization;

namespace EmpleadosMan.Controllers
{
    public class EmpleadosController : Controller
    {
        // GET: Empleados
        public ActionResult Index()
        {
            ViewBag.Lista = "Subir aqui el Excel";
            return View();
        }


        [HttpPost]
        public ActionResult Importar(HttpPostedFileBase excelFile)
        {
            if (excelFile == null || excelFile.ContentLength == 0)
            {
                ViewBag.Error = "Por favor selecciona un archivo Excel";
                return View("Index");
            }
            else
            {
                if (excelFile.FileName.EndsWith("xls") | excelFile.FileName.EndsWith("xlsx"))
                {

                    string path = Server.MapPath("~/Content/" + excelFile.FileName);
                    if (System.IO.File.Exists(path))
                    {
                        System.IO.File.Delete(path);
                    }

                    excelFile.SaveAs(path);


                    //Read Data from excel File.

                    Excel.Application application = new Excel.Application();
                    Excel.Workbook workBook = application.Workbooks.Open(path);
                    Excel.Worksheet hojaUnitech = workBook.Sheets[1];

                    string lastRowE = "E";
                    lastRowE += LastRowPerColumn(5, hojaUnitech).ToString();

                    Excel.Range range = hojaUnitech.Range["C5", lastRowE];


                   List<Empleados> listaEmpleados = new List<Empleados>();


                    for (int i = 1; i <= range.Rows.Count; i++)
                    {

                        Empleados emp = new Empleados();
                        emp.Cont = i;

                        string aPaterno = ((Excel.Range)range.Cells[i, 1]).Text;
                        string aMaterno = ((Excel.Range)range.Cells[i, 2]).Text;
                        string nombre = ((Excel.Range)range.Cells[i, 3]).Text;


                        emp.Nombre = ((aPaterno.ToUpper() + " " + aMaterno.ToUpper() + " " + nombre.ToUpper()));

                        emp.Empresa = ((Excel.Range)range.Cells[i, 5]).Text;
                        emp.FIngreso = ((Excel.Range)range.Cells[i, 6]).Text;
                        emp.AfiliadoIMSS = ((Excel.Range)range.Cells[i, 12]).Text;
                        emp.Estatus = ((Excel.Range)range.Cells[i, 13]).Text;
                        string temp = ((Excel.Range)range.Cells[i, 14]).Text;

                        string[] tokens = temp.Split('$');
                        emp.SDI = decimal.Parse(tokens[1]);
                        emp.NSS = ((Excel.Range)range.Cells[i, 15]).Text;
                        emp.CURP = ((Excel.Range)range.Cells[i, 16]).Text;

                        listaEmpleados.Add(emp);


                    }


                    workBook.Close(path);

                    return View("Empleados", listaEmpleados);

                }
                else
                {
                    ViewBag.Error = "El tipo de archivo es incorrecto";
                    return View("Index");
                }

            }
        }


        //Special thanks to Vitosh Doynov(vitoshacademy.com) helped me solve getting the last row from specific column, phew!
        static int LastRowPerColumn(int column, Excel.Worksheet wks)
        {
            int lastRow = LastRowTotal(wks);
            while (((wks.Cells[lastRow, column]).Text == "") && (lastRow != 1))
            {
                lastRow--;
            }
            return lastRow;
        }

        static int LastRowTotal(Excel.Worksheet wks)
        {
            Excel.Range lastCell = wks.Cells.SpecialCells(Excel.XlCellType.xlCellTypeLastCell, Type.Missing);
            return lastCell.Row;
        }
    }
}
using System;
using System.Collections.Generic;
using System.Collections;
using System.IO;
using System.Linq;
using System.Text;
using System.Xml;
using System.Xml.Linq;
using Excel = Microsoft.Office.Interop.Excel;
using Microsoft.Office.Interop.Excel;

namespace Read_XML
{
    class ReadXML
    {
        struct Books
        {
            public List<string> adc;
            public string author;
            public string subject;
            public int book_id;
        };

        static void Main(string[] args)
        {

            string nameClass = "";
            string filepath_in = "";
            string filepath_out = "";
            string file_out = "";
            try
            {
                nameClass = args[0];
                filepath_in = args[1];
                filepath_out = args[2];
                //file_out = args[3];
            }
            catch
            {
                Console.WriteLine("Please set argument as: \n");
                Console.WriteLine("1st is the name of Class" + ", " + "2nd is the path of database" + ", " + "3rd is the path of output excel file.");
                Console.Read();
            }

            Books Book1;
            Book1.adc.Add("");
            //string filePaths = Path.GetFullPath(filepath_in);
            //Console.WriteLine(filePaths);
            //string[] filePaths = Directory.GetDirectories(filepath_in, "HDC*", SearchOption.AllDirectories);
            string fileMain = "";
            string fileImplementation = "";
            string[] filePaths = Directory.GetFiles(filepath_in, "*.amd", SearchOption.AllDirectories);
            {
                int dem = 0;
                foreach (var a in filePaths)
                {
                    if (a.Split('\\').Last().Equals(nameClass + ".main" + ".amd"))
                    {
                        fileMain = a;
                        dem++;
                        if (dem == 2) break;

                    }
                    if (a.Split('\\').Last().Equals(nameClass + ".implementation" + ".amd"))
                    {
                        fileImplementation = a;
                        dem++;
                        if (dem == 2) break;
                    }
                }
            }

            List<string> listNameotherInputs = new List<string>();
            List<string> listTypeotherInputs = new List<string>();
            List<string> listNameInputs = new List<string>();
            List<string> listTypeInputs = new List<string>();
            List<string> listkkkk = new List<string>();
            List<string> listENUMInputs = new List<string>();
            Hashtable hashLocalVariable = new Hashtable();
            Hashtable hashLocalParameter = new Hashtable();
            Hashtable hashLocalInput = new Hashtable();
            Hashtable hashImportedParameter = new Hashtable();
            Hashtable hashInput = new Hashtable();

            List<string> listNameImportedParam = new List<string>();
            List<string> listTypeImportedParam = new List<string>();
            List<string> listTypVariable_otherInputs = new List<string>();

            List<string> listNameLocalParam = new List<string>();
            List<string> listTypeLocalParam = new List<string>();

            List<string> listNameLocalVari = new List<string>();
            List<string> listTypeLocalVari = new List<string>();

            List<string> list_all_names = new List<string>();
            List<string> list_all_types = new List<string>();



            foreach (var item in XElement.Load(fileMain).Descendants("Elements").Elements("Element"))
            {
                string k = "";
                string d = "";
                var action = item.Attribute("name").Value;

                if (null != item.Element("ElementAttributes"))
                {
                    k = item.Element("ElementAttributes").Attribute("basicModelType").Value;
                    d = item.Element("ElementAttributes").Attribute("modelType").Value;
                    if ((k == "enumeration") && (d == "scalar"))
                    {
                        k = item.Element("ElementAttributes").Element("ScalarType").Element("EnumerationAttributes").Attribute("enumerationName").Value.Split('/').Last();
                        listTypVariable_otherInputs.Add(k);
                    }
                    else if ((k == "cont") && (d == "scalar"))
                    {
                        listTypVariable_otherInputs.Add(k);
                    }
                    else if ((k == "cont") && (d == "oned"))
                    {
                        listTypVariable_otherInputs.Add("Param1D");
                    }
                    else if ((k == "cont") && (d == "twod"))
                    {
                        listTypVariable_otherInputs.Add("Param2D");
                    }
                    else
                    {
                        listTypVariable_otherInputs.Add(k);
                    }

                }

                listNameotherInputs.Add(item.Attribute("name").Value);

                Console.WriteLine("action: {0}", action);
                Console.WriteLine("type of variable: {0}", k);
            }
            int kk = 0;
            string tui = "";
            foreach (var item in XElement.Load(fileMain).Descendants("Elements").Elements("Element").Elements("ElementAttributes"))
            {

                string action = item.Attribute("modelType").Value;
                if (action == "scalar")
                {
                    if (null != item.Element("ScalarType").Element("PrimitiveAttributes"))
                    {
                        tui = item.Element("ScalarType").Element("PrimitiveAttributes").Attribute("scope").Value + " " + item.Element("ScalarType").Element("PrimitiveAttributes").Attribute("kind").Value;

                        listTypeotherInputs.Add(tui);
                        kk++;
                    }

                    else if (null != item.Element("ScalarType").Element("EnumerationAttributes"))
                    {
                        tui = item.Element("ScalarType").Element("EnumerationAttributes").Attribute("scope").Value + " " + item.Element("ScalarType").Element("EnumerationAttributes").Attribute("kind").Value;
                        listTypeotherInputs.Add(tui);
                        kk++;
                    }

                    else
                    {
                        tui = "NULL";
                        listTypeotherInputs.Add(tui);
                    }
                }

                else if (action == "complex")
                {
                    tui = item.Element("ComplexType").Element("ComplexAttributes").Attribute("componentName").Value;
                    listTypeotherInputs.Add(tui);
                }
                else if (action == "oned")
                {
                    tui = item.Element("DimensionalType").Element("PrimitiveAttributes").Attribute("scope").Value + " " + item.Element("DimensionalType").Element("PrimitiveAttributes").Attribute("kind").Value;
                    listTypeotherInputs.Add(tui);
                }
                else if (action == "twod")
                {
                    tui = item.Element("DimensionalType").Element("PrimitiveAttributes").Attribute("scope").Value + " " + item.Element("DimensionalType").Element("PrimitiveAttributes").Attribute("kind").Value;
                    listTypeotherInputs.Add(tui);
                }
                Console.WriteLine("Type: {0}", tui);

            }

            foreach (var item in XElement.Load(fileMain).Descendants("Arguments").Elements("Argument"))
            {
                string action2;
                string nameofEnum = "";
                if (null != item.Element("ElementAttributes"))
                {
                    action2 = item.Element("ElementAttributes").Attribute("basicModelType").Value;
                    if (item.Element("ElementAttributes").Attribute("basicModelType").Value == "enumeration")
                    {
                        foreach (var item2 in XElement.Load(fileMain).Descendants("Arguments").Elements("Argument").Elements("ElementAttributes").Elements("EnumerationAttributes"))
                        {
                            nameofEnum = item2.Attribute("enumerationName").Value;
                        }
                    }
                    else
                    {
                        nameofEnum = null;
                    }
                    //listTypeInputs.Add(item.Element("ElementAttributes").Attribute("basicModelType").Value);
                    listTypeInputs.Add(action2);

                }
                else
                {
                    action2 = "NULL";
                    listTypeInputs.Add(action2);
                }

                // If the element does not have any attributes
                if (!item.Attributes().Any())
                {
                    // Lets skip it
                    continue;
                }

                // Obtain the value of your action attribute - Possible null reference exception here that should be handled
                string action = item.Attribute("name").Value;
                listNameInputs.Add(action);

                //hashInput.Add(action, action2);
                list_all_names.Add(action);
                list_all_types.Add(action2);
                // Do something with your data
                //Console.WriteLine("action: {0}, filename {1}", action, filename);
                Console.WriteLine("Name: {0}", action);
                Console.WriteLine("Type: {0}", action2);
                Console.WriteLine("NameofENUM: {0}", nameofEnum);
            }

            //foreach (var item in XElement.Load(filepath_in).Descendants("EnumerationAttributes"))
            foreach (var item in XElement.Load(fileMain).Descendants("Arguments").Elements("Argument").Elements("ElementAttributes").Elements("ScalarType").Elements("EnumerationAttributes"))
            {
                string action2;
                action2 = item.Attribute("enumerationName").Value;
                listENUMInputs.Add(item.Attribute("enumerationName").Value);
                // If the element does not have any attributes
                if (!item.Attributes().Any())
                {
                    // Lets skip it
                    continue;
                }

                // Obtain the value of your action attribute - Possible null reference exception here that should be handled

                // Do something with your data
                //Console.WriteLine("action: {0}, filename {1}", action, filename);
                Console.WriteLine("Type: {0}", action2);
            }
            for (int i = 0; i < listENUMInputs.Count; i++)
            {
                listENUMInputs[i] = listENUMInputs[i].Split('/').Last();
            }

            for (int i = 0; i < listNameotherInputs.Count; i++)
            {
                if (listTypeotherInputs[i] == "local variable")
                {
                    //hashLocalVariable.Add(listNameotherInputs[i], listTypeotherInputs[i]);
                    listNameLocalVari.Add(listNameotherInputs[i]);
                    listTypeLocalVari.Add(listTypVariable_otherInputs[i]);

                    list_all_names.Add(listNameotherInputs[i]);
                    list_all_types.Add(listTypVariable_otherInputs[i]);
                }
                else if (listTypeotherInputs[i] == "local parameter")
                {
                    //hashLocalParameter.Add(listNameotherInputs[i], listTypeotherInputs[i]);
                    listNameLocalParam.Add(listNameotherInputs[i]);
                    listTypeLocalParam.Add(listTypVariable_otherInputs[i]);

                    list_all_names.Add(listNameotherInputs[i]);
                    list_all_types.Add(listTypVariable_otherInputs[i]);
                }
                else if (listTypeotherInputs[i] == "imported parameter")
                {
                    //hashImportedParameter.Add(listNameotherInputs[i], listTypeotherInputs[i]);
                    listNameImportedParam.Add(listNameotherInputs[i]);
                    listTypeImportedParam.Add(listTypVariable_otherInputs[i]);

                    list_all_names.Add(listNameotherInputs[i]);
                    list_all_types.Add(listTypVariable_otherInputs[i]);
                }
            }

            Excel.Application xlApp;
            Excel.Workbook xlWorkBook;
            Excel.Worksheet xlWorkSheet;
            object misValue = System.Reflection.Missing.Value;

            xlApp = new Excel.Application();
            xlWorkBook = xlApp.Workbooks.Add(misValue);
            xlWorkSheet = (Excel.Worksheet)xlWorkBook.Worksheets.get_Item(1);

            int rowNAME = 10, columnNAME = 3, columnTYPE = 2;

            int ENUMinput = 0;

            xlWorkSheet.Cells[6, 2] = "Tolerance";
            xlWorkSheet.Cells[6, 2].EntireRow.Font.Bold = true;
            xlWorkSheet.Cells[7, columnTYPE] = "Type";
            xlWorkSheet.Cells[7, 2].EntireRow.Font.Bold = true;
            xlWorkSheet.Cells[8, 2] = "Max";
            xlWorkSheet.Cells[8, 2].EntireRow.Font.Bold = true;
            xlWorkSheet.Cells[9, 2] = "Min";
            xlWorkSheet.Cells[9, 2].EntireRow.Font.Bold = true;
            xlWorkSheet.Cells[10, 2] = "TC No.";
            xlWorkSheet.Cells[10, 2].EntireRow.Font.Bold = true;
            xlWorkSheet.Cells[rowNAME, columnNAME].Interior.Color = XlRgbColor.rgbLimeGreen;
            xlWorkSheet.Cells[rowNAME, columnNAME] = "INPUTS";
            xlWorkSheet.Cells[rowNAME, columnNAME].EntireRow.Font.Bold = true;

            rowNAME++;
            for (int i = 0; i < listNameInputs.Count; i++)
            {
                columnTYPE++;
                xlWorkSheet.Cells[rowNAME, columnNAME] = listNameInputs[i];
                if (listTypeInputs[i] == "enumeration")
                {
                    xlWorkSheet.Cells[7, columnTYPE] = listENUMInputs[ENUMinput];
                    ENUMinput++;
                }
                else
                {
                    xlWorkSheet.Cells[7, columnTYPE] = listTypeInputs[i];
                }
                columnNAME++;
            }
            columnTYPE++;

            if (listNameImportedParam != null)
            {
                xlWorkSheet.Cells[--rowNAME, columnNAME].Interior.Color = XlRgbColor.rgbYellow;
                rowNAME++;

                xlWorkSheet.Cells[10, columnNAME] = "IMPORTED PARAMETERS";
                for (int i = 0; i < listNameImportedParam.Count; i++)
                {
                    xlWorkSheet.Cells[11, columnNAME] = listNameImportedParam[i];
                    xlWorkSheet.Cells[7, columnTYPE] = listTypeImportedParam[i];
                    columnTYPE++;
                    columnNAME++;
                }
            }

            if (listNameLocalParam != null)
            {
                xlWorkSheet.Cells[--rowNAME, columnNAME].Interior.Color = XlRgbColor.rgbLightSkyBlue;
                rowNAME++;

                xlWorkSheet.Cells[10, columnNAME] = "LOCAL PARAMETERS";
                for (int i = 0; i < listNameLocalParam.Count; i++)
                {
                    xlWorkSheet.Cells[11, columnNAME] = listNameLocalParam[i];
                    xlWorkSheet.Cells[7, columnTYPE] = listTypeLocalParam[i];
                    columnTYPE++;
                    columnNAME++;
                }
            }

            if (listNameLocalVari != null)
            {
                xlWorkSheet.Cells[--rowNAME, columnNAME].Interior.Color = XlRgbColor.rgbLightGray;
                rowNAME++;

                xlWorkSheet.Cells[10, columnNAME] = "LOCAL VARIABLES";
                for (int i = 0; i < listNameLocalVari.Count; i++)
                {
                    xlWorkSheet.Cells[11, columnNAME] = listNameLocalVari[i];
                    xlWorkSheet.Cells[7, columnTYPE] = listTypeLocalVari[i];
                    columnTYPE++;
                    columnNAME++;
                }
            }
            //int a = listNameInputs.Count + listNameImportedParam.Count + listNameLocalParam.Count + listNameLocalVari.Count;
            //List<string> list_all = new List<string>();
            //string[] g = listNameInputs + listNameImportedParam + listNameLocalParam + listNameLocalVari;
            List<long?> max_physical = new List<long?>();
            List<long?> min_physical = new List<long?>();
            List<double?> tolerance = new List<double?>();

            long max_implement, min_implement;

            foreach (var item in XElement.Load(fileImplementation).Descendants("ImplementationSet").Elements("ImplementationEntry")
                .Elements("ImplementationVariant").Elements("ElementImplementation"))
            {
                for (int i = 0; i < list_all_names.Count; i++)

                    if (item.Attribute("elementName").Value.Equals(list_all_names[i]))
                    {
                        if (list_all_types[i] == "log")
                        {
                            continue;
                        }
                        if (list_all_types[i] == "cont" || list_all_types[i] == "udisc" || list_all_types[i] == "sdisc")
                        {
                            if(null!= item.Element("ScalarImplementation").Element("NumericImplementation").Element("PhysicalInterval"))
                            {
                                string rr = (item.Element("ScalarImplementation").Element("NumericImplementation").Element("PhysicalInterval").Attribute("max").Value);
                                try
                                {
                                    max_physical[i] = long.Parse(rr);
                                    min_physical[i] = long.Parse(item.Element("ScalarImplementation").Element("NumericImplementation").Element("PhysicalInterval").Attribute("min").Value);
                                    //long.TryParse(item.Element("ScalarImplementation").Element("NumericImplementation").Element("PhysicalInterval").Attribute("max").Value, out max_physical[i]);
                                    //long.TryParse(item.Element("ScalarImplementation").Element("NumericImplementation").Element("PhysicalInterval").Attribute("min").Value, out min_physical[i]);

                                    long.TryParse(item.Element("ScalarImplementation").Element("NumericImplementation").Element("ImplementationInterval").Attribute("max").Value, out max_implement);
                                    long.TryParse(item.Element("ScalarImplementation").Element("NumericImplementation").Element("ImplementationInterval").Attribute("min").Value, out min_implement);
                                    tolerance[i] = (max_physical[i] + min_physical[i]) / (max_implement + min_implement);
                                }
                                catch
                                {
                                    //max_physical[i].Value = 0;
                                    //min_physical[i] = 0;
                                    //tolerance[i] = 0;
                                }
                                
                            }
                            else
                            {

                            }
                            
                        }

                        //long.TryParse(item.Element("ScalarImplementation").Element("NumericImplementation").Element("PhysicalInterval").Attribute("max").Value, out max);
                        //action2 = item.Element("ElementAttributes").Attribute("basicModelType").Value;
                    }


            }

            xlWorkSheet.Columns.AutoFit();

            xlWorkSheet.get_Range((Range)xlWorkSheet.Cells[6, 2], (Range)xlWorkSheet.Cells[11, --columnNAME]).Cells.Borders.LineStyle = XlLineStyle.xlContinuous;

            try
            {
                xlWorkBook.SaveAs(filepath_out);
            }

            catch
            {
                //System.Runtime.InteropServices.COMException
                Console.WriteLine("Please close your opening excel file!!!");
                //Console.ReadKey();
                Console.Read();
            }


            xlWorkBook.Close();
            xlApp.Quit();
            System.Runtime.InteropServices.Marshal.ReleaseComObject(xlWorkBook);
            System.Runtime.InteropServices.Marshal.ReleaseComObject(xlWorkSheet);
            System.Runtime.InteropServices.Marshal.ReleaseComObject(xlApp);
            xlWorkBook = null;
            xlWorkSheet = null;
            xlApp = null;
            GC.Collect();
        }
    }

}

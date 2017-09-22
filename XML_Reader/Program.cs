using System;
using Microsoft.Office.Interop.Access.Dao;

namespace XML_Reader
{
    class Program
    {
        static void Main(string[] args)
        {
            //'C:\Users\niraj\Documents\Civictrack\XML_Reader\XML_Reader\bin\Debug\EdmontonManning.xml
            //XmlDataDocument xmldoc = new XmlDataDocument();
            //XmlNodeList xmlnode;
            //int i = 0;
            //string str = null;
            //FileStream fs = new FileStream("EdmontonManning.xml", FileMode.Open, FileAccess.Read);
            //xmldoc.Load(fs);
            //xmlnode = xmldoc.GetElement;sByTagName("Table1");
            //for (i = 0; i <= xmlnode.Count - 1; i++)
            //{
            //    xmlnode[i].ChildNodes.Item(0).InnerText.Trim();
            //    str = xmlnode[i].ChildNodes.Item(0).InnerText.Trim() + "  " + xmlnode[i].ChildNodes.Item(1).InnerText.Trim() + "  " + xmlnode[i].ChildNodes.Item(2).InnerText.Trim();

            //    Console.WriteLine(str);
            //    Console.ReadLine();

            //}

            var dbe = new DBEngine();
            Database db = dbe.OpenDatabase(@"ACCESDATA_BASE_PDF");
            Recordset rstMain = db.OpenRecordset(
                    "SELECT Document FROM Table1 WHERE ID=1017",
                    RecordsetTypeEnum.dbOpenSnapshot);
            Recordset2 rstAttach = rstMain.Fields["Document"].Value;
            Field2 fld = (Field2)rstAttach.Fields["FileData"];
            fld.SaveToFile(@"filename_full_path.pdf");
            //while ((!"Document1.pdf".Equals(rstAttach.Fields["FileName"].Value)) && (!rstAttach.EOF))
            //{
            //    rstAttach.MoveNext();
            //}
            //if (rstAttach.EOF)
            //{
            //    Console.WriteLine("Not found.");
            //}
            //else
            //{
            //    Field2 fld = (Field2)rstAttach.Fields["FileData"];
            //    fld.SaveToFile(@"C:\Users\Gord\Desktop\FromSaveToFile.pdf");
            //}
            db.Close();


            //XmlDocument docu = new XmlDocument();
            //docu.Load("Test.xml");
            //XmlNodeList nodeList = docu.GetElementsByTagName("FileData");
            //string filedata = string.Empty;
            //foreach (XmlNode node in nodeList)
            //{
            //    filedata = node.InnerText;
                

            //    byte[] binaryData = Encoding.UTF8.GetBytes(filedata);
            //    //File.WriteAllBytes(@"C:\textxml.pdf", binaryData);


            //    BinaryWriter writer = new BinaryWriter(File.Open(@"C:\Users\niraj\Documents\Civictrack\textxml1.pdf", FileMode.Create));
            //    writer.Write(filedata);
            //    //string s = Encoding.UTF8.GetString(binaryData);

            //    Console.WriteLine(filedata);
            //    break;
            //}
            Console.ReadLine();

        }
    }
}

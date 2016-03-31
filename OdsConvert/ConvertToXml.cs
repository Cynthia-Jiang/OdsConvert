using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Xml;
using Ionic.Zip;
using System.IO;
using System.Data;
using System.Globalization;

namespace OdsConvert
{
    public class ConvertToXml
    {
        private XmlDocument NewFile;
        private XmlNode NewDoc;
        private XmlNode NewSheet;
        private XmlNode NewRow;
        private XmlNode NewCell;
        public ConvertToXml() { }

        private static string[,] namespaces = new string[,] 
        {
            {"table", "urn:oasis:names:tc:opendocument:xmlns:table:1.0"},
            {"office", "urn:oasis:names:tc:opendocument:xmlns:office:1.0"},
            {"style", "urn:oasis:names:tc:opendocument:xmlns:style:1.0"},
            {"text", "urn:oasis:names:tc:opendocument:xmlns:text:1.0"},            
            {"draw", "urn:oasis:names:tc:opendocument:xmlns:drawing:1.0"},
            {"fo", "urn:oasis:names:tc:opendocument:xmlns:xsl-fo-compatible:1.0"},
            {"dc", "http://purl.org/dc/elements/1.1/"},
            {"meta", "urn:oasis:names:tc:opendocument:xmlns:meta:1.0"},
            {"number", "urn:oasis:names:tc:opendocument:xmlns:datastyle:1.0"},
            {"presentation", "urn:oasis:names:tc:opendocument:xmlns:presentation:1.0"},
            {"svg", "urn:oasis:names:tc:opendocument:xmlns:svg-compatible:1.0"},
            {"chart", "urn:oasis:names:tc:opendocument:xmlns:chart:1.0"},
            {"dr3d", "urn:oasis:names:tc:opendocument:xmlns:dr3d:1.0"},
            {"math", "http://www.w3.org/1998/Math/MathML"},
            {"form", "urn:oasis:names:tc:opendocument:xmlns:form:1.0"},
            {"script", "urn:oasis:names:tc:opendocument:xmlns:script:1.0"},
            {"ooo", "http://openoffice.org/2004/office"},
            {"ooow", "http://openoffice.org/2004/writer"},
            {"oooc", "http://openoffice.org/2004/calc"},
            {"dom", "http://www.w3.org/2001/xml-events"},
            {"xforms", "http://www.w3.org/2002/xforms"},
            {"xsd", "http://www.w3.org/2001/XMLSchema"},
            {"xsi", "http://www.w3.org/2001/XMLSchema-instance"},
            {"rpt", "http://openoffice.org/2005/report"},
            {"of", "urn:oasis:names:tc:opendocument:xmlns:of:1.2"},
            {"rdfa", "http://docs.oasis-open.org/opendocument/meta/rdfa#"},
            {"config", "urn:oasis:names:tc:opendocument:xmlns:config:1.0"}
        };

        private ZipFile GetZipFile(string readFilePath)
        {
            return ZipFile.Read(readFilePath);
        }

        private XmlDocument GetContentXmlFile(ZipFile zipFile)
        {
            // Get file(in zip archive) that contains data ("content.xml").
            ZipEntry contentZipEntry = zipFile["content.xml"];

            // Extract that file to MemoryStream.
            Stream contentStream = new MemoryStream();
            contentZipEntry.Extract(contentStream);
            contentStream.Seek(0, SeekOrigin.Begin);

            // Create XmlDocument from MemoryStream (MemoryStream contains content.xml).
            XmlDocument contentXml = new XmlDocument();
            contentXml.Load(contentStream);

            return contentXml;
        }

        private XmlNamespaceManager InitializeXmlNamespaceManager(XmlDocument xmlDocument)
        {
            XmlNamespaceManager nmsManager = new XmlNamespaceManager(xmlDocument.NameTable);

            for (int i = 0; i < namespaces.GetLength(0); i++)
                nmsManager.AddNamespace(namespaces[i, 0], namespaces[i, 1]);

            return nmsManager;
        }

        public void ReadOdsFile(string inputFilePath)
        {
            NewFile= new XmlDocument();
            //newFile.Name = "odsdoc";

            ZipFile odsZipFile = this.GetZipFile(inputFilePath);

            // Get content.xml file
            XmlDocument contentXml = this.GetContentXmlFile(odsZipFile);

            // Initialize XmlNamespaceManager
            XmlNamespaceManager nmsManager = this.InitializeXmlNamespaceManager(contentXml);

            NewDoc = NewFile.CreateNode("element", "OdsDoc", "");

            foreach (XmlNode tableNode in this.GetTableNodes(contentXml, nmsManager))
            {
                //XmlNode NewSheet = newFile.CreateNode("element","table","");
                NewDoc.AppendChild(this.GetSheet(tableNode,nmsManager));//给根节点添加不同的sheet
            }

            NewFile.AppendChild(NewDoc);

        }

        //write the new file to destination
        public void WriteXmlFile(string outputFilePath)
        {
            NewFile.Save(outputFilePath);
        }

        private XmlNodeList GetTableNodes(XmlDocument contentXmlDocument, XmlNamespaceManager nmsManager)
        {
            return contentXmlDocument.SelectNodes("/office:document-content/office:body/office:spreadsheet/table:table", nmsManager);
        }

        private XmlNode GetSheet(XmlNode tableNode, XmlNamespaceManager nmsManager)
        {
            NewSheet = NewFile.CreateNode("element", "sheet", "");

            XmlNodeList rowNodes = tableNode.SelectNodes("table:table-row", nmsManager);

            int rowIndex = 0;
            foreach (XmlNode rowNode in rowNodes)
            {
                NewSheet.AppendChild(this.GetRow(rowNode, nmsManager, ref rowIndex));
                
            }             
            return NewSheet;
        }

        private XmlNode GetRow(XmlNode rowNode, XmlNamespaceManager nmsManager, ref int rowIndex)
        {
            NewRow = NewFile.CreateNode("element", "row", "");

            XmlAttribute rowsRepeated = rowNode.Attributes["table:number-rows-repeated"];
          //  if (rowsRepeated == null || Convert.ToInt32(rowsRepeated.Value, CultureInfo.InvariantCulture) == 1)
            //{

                XmlNodeList cellNodes = rowNode.SelectNodes("table:table-cell", nmsManager);

                int cellIndex = 1;
                foreach (XmlNode cellNode in cellNodes)
                {
                    NewRow.AppendChild(this.GetCell(cellNode, nmsManager, ref cellIndex));
                    XmlAttribute cellRepeated = cellNode.Attributes["table:number-columns-repeated"];
                    if (cellRepeated != null)
                    {
                        for (int i = 1; i <= Convert.ToInt32(cellRepeated.Value, CultureInfo.InvariantCulture);i++ )
                        {
                            NewCell = NewFile.CreateNode("element", "column" + Convert.ToString(cellIndex), "");
                            NewRow.AppendChild(NewCell);
                            cellIndex++;
                        }
                    }
                }
         /*       rowIndex++;
            }
            else
            {
                rowIndex += Convert.ToInt32(rowsRepeated.Value, CultureInfo.InvariantCulture);
            }*/

            return NewRow;
        }

        private XmlNode GetCell(XmlNode cellNode, XmlNamespaceManager nmsManager, ref int cellIndex)
        {
            NewCell = NewFile.CreateNode("element", "column" + Convert.ToString(cellIndex), "");

            XmlAttribute cellRepeated = cellNode.Attributes["table:number-columns-repeated"];

            NewCell.InnerText = this.ReadCellValue(cellNode);

            cellIndex++;

            return NewCell;
        }

        private string ReadCellValue(XmlNode cell)
        {
            XmlAttribute cellVal = cell.Attributes["office:value"];

            if (cellVal == null)
                return String.IsNullOrEmpty(cell.InnerText) ? null : cell.InnerText;
            else
                return cellVal.Value;
        }
    }
}

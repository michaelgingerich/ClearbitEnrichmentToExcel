/*  ClearbitEnrichmentToExcel
 *  Parse one subschema from multiple JSON responses and with a variable number of properties and values for each response into an Excel worksheet.  
 *  I will add the ability to parse multiple subschemas into multiple Excel worksheets at a later date.
 *  
 *  Copyright(C) 2016  Michael A. Gingerich
 *
 *   This program is free software: you can redistribute it and/or modify
 *   it under the terms of the GNU General Public License as published by
 *   the Free Software Foundation, either version 3 of the License, or
 *   (at your option) any later version.
 *
 *   This program is distributed in the hope that it will be useful,
 *   but WITHOUT ANY WARRANTY; without even the implied warranty of
 *   MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.See the
 *   GNU General Public License for more details.
 *
 *   You should have received a copy of the GNU General Public License
 *   along with this program.If not, see<http://www.gnu.org/licenses/>.
 *
 *   Also add information on how to contact you by electronic and paper mail.
 *
 *   If the program does terminal interaction, make it output a short notice like this when it starts in an interactive mode:
 *
 *   ClearbitEnrichmentToExcel Copyright (C) 2016  Michael A. Gingerich
 *   This program comes with ABSOLUTELY NO WARRANTY; for details type `show w'.
 *   This is free software, and you are welcome to redistribute it
 *   under certain conditions; type `show c' for details.
 */

using Newtonsoft.Json.Linq;
using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;
using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Linq;
using System.Net;
using System.Text;
using System.Threading.Tasks;

namespace ClearbitEnrichmentToExcel
{
    class Program
    {
        String[] emails = { "alex@alexmaccaw.com", "lachy@stripe.com", "harlow@clearbit.com" };
        String[] responses;
        List<JObject> rootSchemas = new List<JObject>();
        List<String> subSchemas = new List<String>();
        List<DataTable> worksheets = new List<DataTable>();

        static void Main(string[] args)
        {
            Program p = new Program();
            p.GetResponses();
            p.MapJSONToTable();
            p.Serialize();
            Console.ReadLine();
        }

        void GetResponses()
        {
            responses = new String[emails.Length];
            for(int i = 0; i < emails.Length; i++)
            {
                WebRequest request = WebRequest.Create("https://person-stream.clearbit.com/v2/combined/find?email=" + emails[i]);
                request.Method = "GET";
                //ADD YOUR CLEARBIT ENRICHMENT API KEY IN THE BELOW STRING
                string userName = "";
                string password = "";
                string credentials = userName + ":" + password;
                request.Headers["Authorization"] = "Basic " + Convert.ToBase64String(Encoding.ASCII.GetBytes(credentials));
                WebResponse response = request.GetResponse();
                Stream dataStream = response.GetResponseStream();
                StreamReader reader = new StreamReader(dataStream);
                responses[i] = reader.ReadToEnd();
            }
        }

        void MapJSONToTable()
        {
            //JObject rootSchema = JObject.Parse(rootSchemas);

            foreach (String response in responses)
            {
                JObject rootSchema = JObject.Parse(response);
                rootSchemas.Add(rootSchema);
                foreach (JProperty subschema in rootSchema.Children().ToList())
                {
                    if (!subSchemas.Contains(subschema.Name))
                    {
                        subSchemas.Add(subschema.Name);
                    }
                }
            }

            List<String> headers = new List<String>();
            int additionalCols = 0;

            foreach (String subSchema in subSchemas)
            {
                foreach (JObject rootSchema in rootSchemas)
                {
                    int headerCol = 0;
                    foreach (JProperty property in rootSchema[subSchema])
                    {
                        if (!(headers.Count > headerCol))
                        {
                            headers.Insert(headerCol, property.Name);
                        }
                        if (!headers[headerCol + additionalCols].Equals(property.Name))
                        {
                            additionalCols++;
                            headers.Insert(headerCol, property.Name);
                        }
                        headerCol++;

                        if (property.Value.HasValues)
                        {
                            if (property.Value is JArray)
                            {
                                foreach (JToken arrayPropertyElement in property.Children().ToList())
                                {
                                    foreach (JValue arrayPropertyElementField in arrayPropertyElement.Children().ToList())
                                    {
                                        if (!(headers.Count > headerCol))
                                        {
                                            headers.Insert(headerCol, String.Empty);
                                        }
                                        if (!headers[headerCol + additionalCols].Equals(String.Empty))
                                        {
                                            additionalCols++;
                                            headers.Insert(headerCol, String.Empty);
                                        }
                                        headerCol++;
                                    }
                                }
                            }
                            else
                            {
                                List<JToken> childInstances = property.Value.Children().ToList();
                                int i = 0;
                                foreach (JProperty childInstance in childInstances)
                                {
                                    if (childInstance.Value is JArray)
                                    {
                                        if (!(headers.Count > headerCol))
                                        {
                                            headers.Insert(headerCol, childInstance.Name);
                                        }
                                        if (!headers[headerCol + additionalCols].Equals(childInstance.Name))
                                        {
                                            additionalCols++;
                                            headers.Insert(headerCol, childInstance.Name);
                                        }
                                        headerCol++;
                                        foreach (JToken arrayChildInstanceElement in childInstance.Children().ToList())
                                        {
                                            List<JToken> arrayChildInstanceElements = arrayChildInstanceElement.Children().Children().ToList();
                                            if (!(arrayChildInstanceElements.Count > 0))
                                            {
                                                int nextChildInstance = i + 1;
                                                for (int j = headerCol; j < headers.Count; j++)
                                                {
                                                    if (headers[j].Equals(((JProperty)childInstances[nextChildInstance]).Name))
                                                    {
                                                        headerCol = j;
                                                        break;
                                                    }
                                                }
                                            }
                                            else
                                            {
                                                foreach (JProperty arrayChildInstanceElementField in arrayChildInstanceElements)
                                                {
                                                    if (!(headers.Count > headerCol))
                                                    {
                                                        headers.Insert(headerCol, arrayChildInstanceElementField.Name);
                                                    }
                                                    if (!headers[headerCol + additionalCols].Equals(arrayChildInstanceElementField.Name))
                                                    {
                                                        additionalCols++;
                                                        headers.Insert(headerCol, arrayChildInstanceElementField.Name);
                                                    }
                                                    headerCol++;
                                                }
                                            }
                                        }
                                    }
                                    else
                                    {
                                        if (!(headers.Count > headerCol))
                                        {
                                            headers.Insert(headerCol, childInstance.Name);
                                        }
                                        if (!headers[headerCol + additionalCols].Equals(childInstance.Name))
                                        {
                                            additionalCols++;
                                            headers.Insert(headerCol, childInstance.Name);
                                        }
                                        headerCol++;
                                    }
                                    i++;
                                }
                            }
                        }
                        else
                        {

                        }
                    }
                }
                //headers = new List<String>();
                break;
            }

            foreach (String subSchema in subSchemas)
            {
                //Loop through matching subschemas for all responses at this scope.

                DataTable currentWkSht = new DataTable();
                List<DataRow> currentDRows = new List<DataRow>();
                DataRow oldHeaders = currentWkSht.NewRow();
                DataRow data = currentWkSht.NewRow();
                Console.WriteLine(subSchema);
                //worksheet name

                int headerCol = 0, dataCol;
                bool headersExist = false;
                foreach (JObject rootSchema in rootSchemas)
                {
                    dataCol = 0;
                    foreach (JProperty property in rootSchema[subSchema])
                    {
                        Console.WriteLine("\t" + property.Name);
                        //headers
                        if (!headersExist)
                        {
                            currentWkSht.Columns.Add();
                            //oldHeaders[headerCol++] = property.Name;
                        }

                        if (property.Value.HasValues)
                        {
                            dataCol++;
                            if (property.Value is JArray)
                            {
                                foreach (JToken arrayPropertyElement in property.Children().ToList())
                                {
                                    foreach (JValue arrayPropertyElementField in arrayPropertyElement.Children().ToList())
                                    {
                                        Console.WriteLine("\t\t" + arrayPropertyElementField.Value);
                                        //data
                                        if (!headersExist)
                                        {
                                            currentWkSht.Columns.Add();
                                            //    headerCol++;
                                        }
                                        data[dataCol++] = arrayPropertyElementField.Value;
                                    }
                                }
                            }
                            else
                            {
                                List<JToken> childInstances = property.Value.Children().ToList();
                                int i = 0;
                                foreach (JProperty childInstance in childInstances)
                                {
                                    if (childInstance.Value is JArray)
                                    {
                                        Console.WriteLine("\t\t" + childInstance.Name);
                                        //headers
                                        if (!headersExist)
                                        {
                                            currentWkSht.Columns.Add();
                                            //    oldHeaders[headerCol++] = childInstance.Name;
                                        }
                                        dataCol++;
                                        foreach (JToken arrayChildInstanceElement in childInstance.Children().ToList())
                                        {
                                            List<JToken> arrayChildInstanceElements = arrayChildInstanceElement.Children().Children().ToList();
                                            if (!(arrayChildInstanceElements.Count > 0))
                                            {
                                                int nextChildInstance = i + 1;
                                                for (int j = dataCol; j < headers.Count; j++)
                                                {
                                                    if (!headersExist)
                                                    {
                                                        currentWkSht.Columns.Add();
                                                    }
                                                    data[dataCol++] = String.Empty;
                                                    if (headers[j].Equals(((JProperty)childInstances[nextChildInstance]).Name))
                                                    {
                                                        dataCol = j;
                                                        break;
                                                    }
                                                }
                                            }
                                            else
                                            {
                                                foreach (JProperty arrayChildInstanceElementField in arrayChildInstanceElement.Children().Children().ToList())
                                                {
                                                    Console.WriteLine("\t\t\t" + arrayChildInstanceElementField.Name);
                                                    //header
                                                    if (!headersExist)
                                                    {
                                                        currentWkSht.Columns.Add();
                                                        //oldHeaders[headerCol++] = arrayChildInstanceElementField.Name;
                                                    }
                                                    Console.WriteLine("\t\t\t\t" + arrayChildInstanceElementField.Value);
                                                    //data
                                                    data[dataCol++] = arrayChildInstanceElementField.Value;
                                                }
                                            }
                                        }
                                    }
                                    else
                                    {
                                        Console.WriteLine("\t\t" + childInstance.Name);
                                        //header
                                        if (!headersExist)
                                        {
                                            currentWkSht.Columns.Add();
                                            //    oldHeaders[headerCol++] = childInstance.Name;
                                        }
                                        Console.WriteLine("\t\t\t" + childInstance.Value);
                                        //data
                                        //METRICS
                                        data[dataCol++] = childInstance.Value;
                                    }
                                    i++;
                                }
                            }
                        }
                        else
                        {
                            Console.WriteLine("\t\t" + property.Value);
                            //data
                            data[dataCol++] = property.Value;
                        }
                    }
                    if (!headersExist)
                    {
                        if (headers.Count > currentWkSht.Columns.Count)
                        {
                            for (int i = currentWkSht.Columns.Count; i < headers.Count; i++)
                            {
                                currentWkSht.Columns.Add();
                            }
                        }
                        currentWkSht.Rows.Add(headers.ToArray());
                        headersExist = true;
                    }
                    currentWkSht.Rows.Add(data);
                    data = currentWkSht.NewRow();
                }
                worksheets.Add(currentWkSht);
                break;
            }
        }

        void Serialize()
        {
            using (FileStream stream = new FileStream(@"Responses.xlsx", FileMode.Create, FileAccess.Write))
            {
                IWorkbook wb = new XSSFWorkbook();
                ICreationHelper cH = wb.GetCreationHelper();
                int k = 0;
                foreach (String subSchema in subSchemas)
                {
                    ISheet sheet = wb.CreateSheet(subSchema);

                    for (int i = 0; i < worksheets[k].Rows.Count; i++)
                    {
                        IRow row = sheet.CreateRow(i);

                        int j = 0;
                        foreach (Object cellVal in worksheets[k].Rows[i].ItemArray)
                        {
                            ICell cell = row.CreateCell(j);
                            cell.SetCellValue(cH.CreateRichTextString(worksheets[k].Rows[i].ItemArray[j].ToString()));
                            j++;
                        }
                    }
                    k++;
                    break;
                }
                wb.Write(stream);
            }
        }
    }
}

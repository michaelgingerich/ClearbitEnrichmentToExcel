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
        }

        void GetResponses()
        {
            responses = new String[emails.Length];
            for(int i = 0; i < emails.Length; i++)
            {
                WebRequest request = WebRequest.Create("https://person-stream.clearbit.com/v2/combined/find?email=" + emails[i]);
                request.Method = "GET";
                string userName = "YOUR API KEY";
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

            List<List<String>> subSchemaHeaders = new List<List<String>>();
            int subSchemaIndex = 0;

            foreach (String subSchema in subSchemas)
            {
                subSchemaHeaders.Add(new List<String>());
                foreach (JObject rootSchema in rootSchemas)
                {
                    int headerCol = 0;
                    foreach (JProperty property in rootSchema[subSchema])
                    {
                        //Insert property name
                        if (!(subSchemaHeaders[subSchemaIndex].Count > headerCol))
                        {
                            subSchemaHeaders[subSchemaIndex].Insert(headerCol, property.Name);
                        }
                        if (!subSchemaHeaders[subSchemaIndex][headerCol].Equals(property.Name))
                        {
                            subSchemaHeaders[subSchemaIndex].Insert(headerCol, property.Name);
                        }
                        headerCol++;

                        if (property.Value.HasValues)
                        {
                            if (property.Value is JArray)
                            {
                                List<JToken> arrayPropertyElements = property.Children().Children().ToList();
                                var myVar = arrayPropertyElements[0];
                                //Get the headerCol that occurs after this JArray in this rootSchema.
                                int j = arrayPropertyElements.Count + headerCol;
                                List<JToken> properties = rootSchema[subSchema].Children().ToList();
                                //Get index for next property.
                                int nextProperty = 1 + properties.IndexOf(property);
                                //Get the headerCol that occurs after this JArray in subSchemaHeaders[subSchemaIndex].
                                for (int k = headerCol; k < subSchemaHeaders[subSchemaIndex].Count; k++)
                                {
                                    try
                                    {
                                        //Check for String.Empty instead.
                                        if (subSchemaHeaders[subSchemaIndex][k].Equals(((JProperty)properties[nextProperty]).Name))
                                        {
                                            headerCol = k;
                                            break;
                                        }
                                    }
                                    catch (ArgumentOutOfRangeException e1)
                                    {
                                        headerCol = subSchemaHeaders[subSchemaIndex].Count;
                                        break;
                                    }
                                }
                                //Check that this JArray requires more spaces than the spaces accounted for in subSchemaHeaders[subSchemaIndex] and insert spaces accordingly.
                                if (j > headerCol)
                                {
                                    for (int l = headerCol; l < j; l++)
                                    {
                                        subSchemaHeaders[subSchemaIndex].Insert(l, String.Empty);
                                    }
                                    headerCol = j;
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
                                        //Insert child-instance name
                                        if (!(subSchemaHeaders[subSchemaIndex].Count > headerCol))
                                        {
                                            subSchemaHeaders[subSchemaIndex].Insert(headerCol, childInstance.Name);
                                        }
                                        if (!subSchemaHeaders[subSchemaIndex][headerCol].Equals(childInstance.Name))
                                        {
                                            subSchemaHeaders[subSchemaIndex].Insert(headerCol, childInstance.Name);
                                        }
                                        headerCol++;
                                        foreach (JToken arrayChildInstanceElement in childInstance.Children().ToList())
                                        {
                                            List<JToken> childInstanceArrayElements = arrayChildInstanceElement.Children().ToList();
                                            List<JToken> childChildInstanceArrayElements = arrayChildInstanceElement.Children().Children().ToList();
                                            //Check if this JArray has any elements at all.
                                            if (!(childInstanceArrayElements.Count > 0))
                                            {
                                                //Get the headerCol for the next childInstance.
                                                int nextChildInstance = i + 1;
                                                for (int j = headerCol; j < subSchemaHeaders[subSchemaIndex].Count; j++)
                                                {
                                                    //Check for String.Empty
                                                    if (subSchemaHeaders[subSchemaIndex][j].Equals(((JProperty)childInstances[nextChildInstance]).Name))
                                                    {
                                                        headerCol = j;
                                                        break;
                                                    }
                                                }
                                            }
                                            else
                                            {
                                                //Check if this JArray is not nested in an object.
                                                if (!(childChildInstanceArrayElements.Count > 0))
                                                {
                                                    //Get the headerCol that occurs after this JArray in this rootSchema.
                                                    int j = childInstanceArrayElements.Count + headerCol;
                                                    //Get index for next childInstance.
                                                    int nextChildInstance = i + 1;
                                                    int k = headerCol;
                                                    //Check that this is not the last childInstance for this property.
                                                    if (childInstances.Count > nextChildInstance)
                                                    {
                                                        //Get the headerCol that occurs after this JArray in subSchemaHeaders[subSchemaIndex].
                                                        for (; k < subSchemaHeaders[subSchemaIndex].Count; k++)
                                                        {
                                                            if (subSchemaHeaders[subSchemaIndex][k].Equals(((JProperty)childInstances[nextChildInstance]).Name))
                                                            {
                                                                headerCol = k;
                                                                break;
                                                            }
                                                        }
                                                    }
                                                    else
                                                    {
                                                        List<JToken> properties = rootSchema[subSchema].Children().ToList();
                                                        int nextProperty = 1 + properties.IndexOf(property);
                                                        //Get the headerCol for the next property.
                                                        for (; k < subSchemaHeaders[subSchemaIndex].Count; k++)
                                                        {
                                                            if (subSchemaHeaders[subSchemaIndex][k].Equals(((JProperty)properties[nextProperty]).Name))
                                                            {
                                                                headerCol = k;
                                                                break;
                                                            }
                                                        }
                                                    }
                                                    //Check that this JArray requires more spaces than the spaces accounted for in subSchemaHeaders[subSchemaIndex] and insert spaces accordingly.
                                                    if (j > headerCol)
                                                    {
                                                        for (int l = headerCol; l < j; l++)
                                                        {
                                                            subSchemaHeaders[subSchemaIndex].Insert(l, String.Empty);
                                                        }
                                                        headerCol = j;
                                                    }
                                                }
                                                else
                                                {
                                                    foreach (JProperty childChildInstanceArrayElementField in childChildInstanceArrayElements)
                                                    {
                                                        //Insert element name.
                                                        if (!(subSchemaHeaders[subSchemaIndex].Count > headerCol))
                                                        {
                                                            subSchemaHeaders[subSchemaIndex].Insert(headerCol, childChildInstanceArrayElementField.Name);
                                                        }
                                                        if (!subSchemaHeaders[subSchemaIndex][headerCol].Equals(childChildInstanceArrayElementField.Name))
                                                        {
                                                            subSchemaHeaders[subSchemaIndex].Insert(headerCol, childChildInstanceArrayElementField.Name);
                                                        }
                                                        headerCol++;
                                                    }
                                                }
                                            }
                                        }
                                    }
                                    else
                                    {
                                        //Insert child-instance name
                                        if (!(subSchemaHeaders[subSchemaIndex].Count > headerCol))
                                        {
                                            subSchemaHeaders[subSchemaIndex].Insert(headerCol, childInstance.Name);
                                        }
                                        if (!subSchemaHeaders[subSchemaIndex][headerCol].Equals(childInstance.Name))
                                        {
                                            subSchemaHeaders[subSchemaIndex].Insert(headerCol, childInstance.Name);
                                        }
                                        headerCol++;
                                    }
                                    i++;
                                }
                            }
                        }
                        else
                        {
                            if (property.Value is JArray)
                            {
                                List<JToken> properties = rootSchema[subSchema].Children().ToList();
                                //Get index for next property.
                                int nextProperty = 1 + properties.IndexOf(property);
                                int k = headerCol;
                                //Get the headerCol that occurs after this JArray in subSchemaHeaders[subSchemaIndex].
                                for (; k < subSchemaHeaders[subSchemaIndex].Count; k++)
                                {
                                    try
                                    {
                                        //Check for String.Empty instead.
                                        if (subSchemaHeaders[subSchemaIndex][k].Equals(((JProperty)properties[nextProperty]).Name))
                                        {
                                            headerCol = k;
                                            break;
                                        }
                                    }
                                    catch (ArgumentOutOfRangeException e1)
                                    {
                                        headerCol = subSchemaHeaders[subSchemaIndex].Count;
                                        break;
                                    }
                                }
                            }
                        }
                    }
                }
                subSchemaIndex++;
            }

            subSchemaIndex = 0;
            foreach (String subSchema in subSchemas)
            {
                DataTable currentWkSht = new DataTable();
                List<DataRow> currentDRows = new List<DataRow>();
                DataRow data = currentWkSht.NewRow();

                int dataCol;
                bool headersExist = false;
                foreach (JObject rootSchema in rootSchemas)
                {
                    dataCol = 0;
                    foreach (JProperty property in rootSchema[subSchema])
                    {
                        if (!headersExist)
                        {
                            currentWkSht.Columns.Add();
                        }
                        if (property.Value.HasValues)
                        {
                            dataCol++;
                            if (property.Value is JArray)
                            {
                                List<JToken> arrayPropertyElements = property.Children().Children().ToList();
                                //Get the headerCol that occurs after this JArray in this rootSchema.
                                int j = arrayPropertyElements.Count + dataCol;
                                List<JToken> properties = rootSchema[subSchema].Children().ToList();
                                //Get index for next property.
                                int nextProperty = 1 + properties.IndexOf(property);
                                //Get the headerCol that occurs after this JArray in data.
                                int k = dataCol;
                                for (; k < subSchemaHeaders[subSchemaIndex].Count; k++)
                                {
                                    try
                                    {
                                        //Check for String.Empty instead.
                                        if (subSchemaHeaders[subSchemaIndex][k].Equals(((JProperty)properties[nextProperty]).Name))
                                        {
                                            break;
                                        }
                                    }
                                    catch (ArgumentOutOfRangeException e1)
                                    {
                                        //dataCol = data.Table.Columns.Count;
                                        k = data.Table.Columns.Count;
                                        break;
                                    }
                                }
                                //Check that this JArray requires more spaces than the spaces accounted for in subSchemaHeaders[subSchemaIndex] and insert spaces accordingly.
                                if (j > dataCol)
                                {
                                    for (int l = dataCol, m = 0; l < j; l++, m++)
                                    {
                                        if (!headersExist)
                                        {
                                            currentWkSht.Columns.Add();
                                        }
                                        data[l] = arrayPropertyElements[m];
                                    }
                                    if (!headersExist)
                                    {
                                        for (int n = j; n < k; n++)
                                        {
                                            currentWkSht.Columns.Add();
                                        }
                                    }
                                    if (j > k)
                                    {
                                        dataCol = j;
                                    }
                                    else
                                    {
                                        dataCol = k;
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
                                        if (!headersExist)
                                        {
                                            currentWkSht.Columns.Add();
                                        }
                                        dataCol++;
                                        foreach (JToken arrayChildInstanceElement in childInstance.Children().ToList())
                                        {
                                            List<JToken> childInstanceArrayElements = arrayChildInstanceElement.Children().ToList();
                                            List<JToken> childChildInstanceArrayElements = arrayChildInstanceElement.Children().Children().ToList();
                                            //Check if this JArray has any elements at all.
                                            if (!(childInstanceArrayElements.Count > 0))
                                            {
                                                int nextChildInstance = i + 1;
                                                for (int j = dataCol; j < subSchemaHeaders[subSchemaIndex].Count; j++)
                                                {
                                                    if (subSchemaHeaders[subSchemaIndex][j].Equals(((JProperty)childInstances[nextChildInstance]).Name))
                                                    {
                                                        dataCol = j;
                                                        break;
                                                    }
                                                    if (!headersExist)
                                                    {
                                                        currentWkSht.Columns.Add();
                                                    }
                                                    data[dataCol++] = String.Empty;
                                                }
                                            }
                                            else
                                            {
                                                //Check if this JArray is not nested in an object.
                                                if (!(childChildInstanceArrayElements.Count > 0))
                                                {
                                                    //Get the headerCol that occurs after this JArray in this rootSchema.
                                                    int j = childInstanceArrayElements.Count + dataCol;
                                                    //Get index for next childInstance.
                                                    int nextChildInstance = i + 1;
                                                    int k = dataCol;
                                                    //Check that this is not the last childInstance for this property.
                                                    if (childInstances.Count > nextChildInstance)
                                                    {
                                                        //Get the headerCol that occurs after this JArray in data.
                                                        for (; k < data.Table.Columns.Count; k++)
                                                        {
                                                            if (subSchemaHeaders[subSchemaIndex][k].Equals(((JProperty)childInstances[nextChildInstance]).Name))
                                                            {
                                                                break;
                                                            }
                                                        }
                                                    }
                                                    else
                                                    {
                                                        List<JToken> properties = rootSchema[subSchema].Children().ToList();
                                                        int nextProperty = 1 + properties.IndexOf(property);
                                                        //Get the headerCol for the next property.
                                                        for (; k < subSchemaHeaders[subSchemaIndex].Count; k++)
                                                        {
                                                            if (subSchemaHeaders[subSchemaIndex][k].Equals(((JProperty)properties[nextProperty]).Name))
                                                            {
                                                                //dataCol = k;
                                                                break;
                                                            }
                                                        }
                                                    }
                                                    //Check that this JArray requires more spaces than the spaces accounted for in subSchemaHeaders[subSchemaIndex] and insert spaces accordingly.
                                                    if (j > dataCol)
                                                    {
                                                        for (int l = dataCol, m = 0; l < j; l++, m++)
                                                        {
                                                            if (!headersExist)
                                                            {
                                                                currentWkSht.Columns.Add();
                                                            }
                                                            data[l] = childInstanceArrayElements[m];
                                                        }
                                                        if (!headersExist)
                                                        {
                                                            for (int n = j; n < k; n++)
                                                            {
                                                                currentWkSht.Columns.Add();
                                                            }
                                                        }
                                                        if (k < j)
                                                        {
                                                            dataCol = j;
                                                        }
                                                        else
                                                        {
                                                            dataCol = k;
                                                        }
                                                    }
                                                }
                                                else
                                                {
                                                    //Get the headerCol that occurs after this JArray in this rootSchema.
                                                    int j = childChildInstanceArrayElements.Count + dataCol;
                                                    //Check that this JArray requires more spaces than the spaces accounted for in subSchemaHeaders[subSchemaIndex] and insert spaces accordingly.
                                                    if (j > dataCol)
                                                    {
                                                        int l = dataCol;
                                                        foreach (JProperty childChildInstanceArrayElement in childChildInstanceArrayElements)
                                                        {
                                                            if (!headersExist)
                                                            {
                                                                currentWkSht.Columns.Add();
                                                            }
                                                            data[l++] = childChildInstanceArrayElement.Value;
                                                        }
                                                        dataCol = j;
                                                    }
                                                }
                                            }
                                        }
                                    }
                                    else
                                    {
                                        if (property.Value is JArray)
                                        {
                                            List<JToken> properties = rootSchema[subSchema].Children().ToList();
                                            //Get index for next property.
                                            int nextProperty = 1 + properties.IndexOf(property);
                                            int k = dataCol;
                                            //Get the headerCol that occurs after this JArray in data.
                                            for (; k < data.Table.Columns.Count; k++)
                                            {
                                                try
                                                {
                                                    //Check for String.Empty instead.
                                                    if (data[k].Equals(((JProperty)properties[nextProperty]).Name))
                                                    {
                                                        dataCol = k;
                                                        break;
                                                    }
                                                }
                                                catch (ArgumentOutOfRangeException e1)
                                                {
                                                    dataCol = data.Table.Columns.Count;
                                                    break;
                                                }
                                            }
                                        }
                                        else
                                        {
                                            if (!headersExist)
                                            {
                                                currentWkSht.Columns.Add();
                                            }
                                            data[dataCol++] = childInstance.Value;
                                        }
                                    }
                                    i++;
                                }
                            }
                        }
                        else
                        {
                            if (property.Value is JArray)
                            {
                                List<JToken> properties = rootSchema[subSchema].Children().ToList();
                                //Get index for next property.
                                int nextProperty = 1 + properties.IndexOf(property);
                                int k = dataCol;
                                //Get the headerCol that occurs after this JArray in subSchemaHeaders[subSchemaIndex].
                                for (; k < subSchemaHeaders[subSchemaIndex].Count; k++)
                                {
                                    try
                                    {
                                        //Check for String.Empty instead.
                                        if (subSchemaHeaders[subSchemaIndex][k].Equals(((JProperty)properties[nextProperty]).Name))
                                        {
                                            break;
                                        }
                                    }
                                    catch (ArgumentOutOfRangeException e1)
                                    {
                                        dataCol = subSchemaHeaders[subSchemaIndex].Count;
                                        break;
                                    }
                                }
                                if (!headersExist)
                                {
                                    for (int l = dataCol; l < k - 1; l++)
                                    {
                                        currentWkSht.Columns.Add();
                                    }
                                }
                                dataCol = k;
                            }
                            else
                            {
                                data[dataCol++] = property.Value;
                            }
                        }
                    }
                    if (!headersExist)
                    {
                        if (subSchemaHeaders[subSchemaIndex].Count > currentWkSht.Columns.Count)
                        {
                            for (int i = currentWkSht.Columns.Count; i < subSchemaHeaders[subSchemaIndex].Count; i++)
                            {
                                currentWkSht.Columns.Add();
                            }
                        }
                        currentWkSht.Rows.Add(subSchemaHeaders[subSchemaIndex].ToArray());
                        headersExist = true;
                    }
                    currentWkSht.Rows.Add(data);
                    data = currentWkSht.NewRow();
                }
                worksheets.Add(currentWkSht);
                subSchemaIndex++;
            }
        }

        void Serialize()
        {
            using (FileStream stream = new FileStream(@"Cust_Data.xlsx", FileMode.Create, FileAccess.Write))
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
                }
                wb.Write(stream);
            }
        }
    }
}

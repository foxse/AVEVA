using Aveva.Core.PMLNet;
using System.Data;
using DataTable = System.Data.DataTable;
using System;
using System.Threading.Tasks;
using Aveva.Core.Database;
using Aveva.Core.Utilities.CommandLine;
using System.Data.OleDb;
using System.Linq;
using System.Windows.Forms;

namespace Aveva.Core.InstLoader
{
    [PMLNetCallable()]

    public class InstLoader
    {
        public DataTable LoadedData { get; set; }
        public DataTable FailedData { get; set; }

        [PMLNetCallable()]
        public InstLoader()
        {

        }

        [PMLNetCallable()]
        public void Assign(InstLoader that)
        {

        }

        [PMLNetCallable()]
        public async void Start()
        {
            var form = new MainWindow();

            form.openFileDialog.ShowDialog();

            var loadedData = ReadFile(form.openFileDialog.FileName);

            ParseData(loadedData);

            var result = await Task.Run(() =>
            {
                form.dataGridView1.DataSource = LoadedData;
                form.TopLevel = true;
                form.TopMost = true;
                form.ShowDialog();
                return form.DialogResult;
            });

            if (result == DialogResult.OK)
            {
                CreateElements();
            }

        }

        private void CreateElements()
        {
            try
            {
                var element = DbElement.GetElement("/INST-GROUP");

                var elemType = DbElementType.GetElementType(":Instrument");

                var sysClassAtt = DbAttribute.GetDbAttribute("SystemClass");
                var tagDescAtt = DbAttribute.GetDbAttribute("TagIdentifierDesc");

                string currentTag = "";
                DbElement currentElement = null;

                int created = 0;

                foreach (DataRow item in LoadedData.Rows)
                {
                    var tagName = item.ItemArray[0].ToString();

                    if (!tagName.StartsWith("/"))
                        tagName = $"/{tagName}";

                    if (tagName != currentTag)
                    {
                        var exists = DbElement.GetElement(tagName);

                        if (exists != null)
                        {
                            Command.CreateCommand($"$P Element {tagName} already exists! Skipped.").Run();
                            continue;
                        }

                        currentElement = element.CreateLast(elemType);

                        if (!DbElement.IsValidNameFormat(tagName, Database.DbType.Engineering, out var message))
                        {
                            var messageText = message.MessageText();
                            Command.CreateCommand($"$P {messageText}").Run();
                        }

                        if (!currentElement.SetAttributeValue(DbAttributeInstance.NAME, tagName))
                        {
                            Command.CreateCommand($"$P Can't set Name {tagName} for element {currentElement.Name()}").Run();

                            Command.CreateCommand($"{currentElement.Name()}").Run();
                            Command.CreateCommand($"set Name {tagName}").Run();
                        }

                        if (tagDescAtt != null)
                            currentElement.SetAttributeValue(tagDescAtt, item.ItemArray[1].ToString());

                        if (sysClassAtt != null)
                            currentElement.SetAttributeValue(sysClassAtt, item.ItemArray[2].ToString());

                        currentTag = tagName;

                        created++;
                    }

                    if (currentElement != null)
                    {
                        var attName = item.ItemArray[3].ToString();
                        var attribute = DbAttribute.GetDbAttribute($":{attName}");

                        if (attribute == null)
							attribute = DbAttribute.GetDbAttribute(attName);

						if (attribute != null)
                        {
                            var attValue = item.ItemArray[4]?.ToString();

                            if (attValue != null && !currentElement.SetAttributeValue(attribute, attValue))
                            {
                                Command.CreateCommand($"$P Can't set {attName} attribute value: '{attValue}' for element {currentElement.Name()}").Run();
                            }
                        }
                    }

                }

                Command.CreateCommand($"$P 'Done! {created} elements was created'").Run();
            }
            catch (Exception ex)
            {
                Command.CreateCommand($"$P '{ex.Message}'").Run();
            }
        }



        private static Tuple<DataTable, DataTable> ReadFile(string path)
        {
            // Microsoft Access Database Engine 2010 Redistributable
            // https://www.microsoft.com/en-us/download/details.aspx?id=13255

            string conn = path.EndsWith(".xls")
                ? @"Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" + path + ";Extended Properties='Excel 8.0;HRD=Yes';"
                : @"Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" + path + ";Extended Properties='Excel 12.0;HDR=Yes';";

            var dt = new DataTable();
            var dt2 = new DataTable();
            using (var oleDb = new OleDbConnection(conn))
            {
                var oleAdpt = new OleDbDataAdapter("select * from [data$]", oleDb);
                oleAdpt.Fill(dt);

                var oleAdpt2 = new OleDbDataAdapter("select * from [mapping$]", oleDb);
                oleAdpt2.Fill(dt2);
            }

            return new Tuple<DataTable, DataTable>(dt, dt2);
        }


        public void ParseData(Tuple<DataTable, DataTable> dataTables)
        {

            try
            {
                dataTables.Item1.Columns.Add("Type");

                var data = new DataTable();

                foreach (DataColumn column in dataTables.Item1.Columns)
                {
                    data.Columns.Add(column.ColumnName);
                }

                dataTables.Item1.AsEnumerable()
                        .Where(i =>
                            i["System Class"].ToString() == "Production"
                            && i["TagIdentifier Desc"].ToString() == "Pressure Gauge")
                        .Join(dataTables.Item2.AsEnumerable(),
						    d => d["TagType"],
							m => m["TagType"],
                            (d, m) => new { d, m })
                        .Select(i => i)
                        .ToList()
                        .ForEach(i =>
                        {
                            i.d["Type"] = i.m["Type"];
                            data.Rows.Add(i.d.ItemArray);
                        });

                LoadedData = new DataTable();

                LoadedData.Columns.Add("Tag");
                LoadedData.Columns.Add("TagIdentifier Desc");
                LoadedData.Columns.Add("System Class");
                LoadedData.Columns.Add("Attribute Name");
                LoadedData.Columns.Add("Attribute Value");

                var columnsToSkip = new string[] { "Tag", "TagIdentifier Desc", "System Class" };

                foreach (DataRow dataRow in data.Rows)
                {
                    var rowData = new string[5];

                    rowData[0] = dataRow["Tag"].ToString();
                    rowData[1] = dataRow["TagIdentifier Desc"].ToString();
                    rowData[2] = dataRow["System Class"].ToString();

                    foreach (DataColumn dataColumn in data.Columns)
                    {
                        if (columnsToSkip.Contains(dataColumn.ColumnName))
                            continue;

                        string cellVal = dataRow[dataColumn.ColumnName].ToString();

                        rowData[3] = dataColumn.ColumnName.Replace(" ", string.Empty);
                        rowData[4] = cellVal;

                        LoadedData.Rows.Add(rowData);
                    }
                }
            }
            catch (Exception ex)
            {
                Command.CreateCommand($"$P {ex.Message}").Run();
            }
        }
    }
}

using System.Collections.Generic;
using System.Data.OleDb;
using System.Data;
using System;
using System.Linq;
using System.IO;

namespace NamingChecker
{
	internal class Program
	{
		private static Dictionary<string, HashSet<string>> _validValues =
			new Dictionary<string, HashSet<string>> {
				{"Wbs Code", new HashSet<string> { "2125", "2137", "2152", "2101", "2136", "2613", "2615", "2714", "2715", "2131", "2141", "2144", "2146", "2650", "2711", "2102", "2103", "2104", "2105", "2106", "2107", "2108", "2124", "2128" } },
				{"Tag Class Name", new HashSet<string> { "ADSORBER", "ADSORBER SET", "BLOWER", "COMPRESSOR", "COMPRESSOR SET", "ELECTRIC MOTOR", "FILTER", "HEAT EXCHANGER SECTION", "ON-OFF VALVE", "PUMP", "PUMP SET", "PLATE HEAT EXCHANGER", "HEAT EXCHANGER SET", "STEAM BOILER PACKAGE", "LINE" } },
				{"Tag Status", new HashSet<string> { "As built", "Commissioned", "Defined", "Deleted", "Delivered", "Dispatched", "Installed", "Purchased", "Ready for shipment" } },
			};

		private static HashSet<string> _equipmentTypes = new HashSet<string> { 
			"A", "BL", "K", "P", "F", "E", "PK"
        };

		private static HashSet<string> _driveTypes = new HashSet<string> {
			"D", "GT", "M", "T"
		};

		private static HashSet<string> _requiredFields =
			new HashSet<string> { "Tag number", "Tag Class Name", "Wbs Code", "Tag Description Eng", "Designed By Company Name", "System code", "Tag status" };

		private static List<string> _allErrors;
		private static HashSet<string> _validTags;


		static void Main(string[] args)
		{
			var testFile = @"C:\Users\User\Downloads\task1_dataRegister_origin.xlsx";

            var dt = ReadFile(args.Length == 0? testFile : args[0]);

			_allErrors = new List<string>();
			_validTags = new HashSet<string>();

            _allErrors.Add("№;Tag number;Tag Class Name;Property name;Error description");

			try
			{
				CheckData(dt);
				ChackParents(dt);

				File.WriteAllLines(testFile.Replace(".xlsx", "_errors.csv").Replace(".xls", "_errors.csv"), 
					_allErrors, System.Text.Encoding.UTF8);
			}
			catch (Exception ex)
			{
				throw;
			}
		}

		private static void ChackParents(DataTable dt)
		{
			foreach (DataRow row in dt.Rows)
			{
				var parentTag = row["Parent Tag Number"].ToString();

				if (parentTag != String.Empty && !_validTags.Contains(parentTag))
					AddErrorMessage(row, "Parent Tag Number", $"Родительский элемет {parentTag} не найден в реестре или имеет ошибки");
			}
		}

		private static void CheckData(DataTable dt)
		{
			foreach (DataRow row in dt.Rows)
			{
				var isValid = true;
				string tag = string.Empty;

				foreach (DataColumn column in dt.Columns)
				{
					string value = row[column].ToString();

					if (!CheckRequredFieldIsFilled(column.ColumnName, value, row))
					{
						isValid = false;
					}
					else if (!CheckFieldValueIsValid(column.ColumnName, value, row))
					{
						isValid = false;
					}

					if (column.ColumnName == "Tag number")
					{
						tag = value;
						if (!CheckTagIsValid(value, row))
							isValid = false;
					}
				}

				if (isValid)
					_validTags.Add(tag);
			}
		}


		private static bool CheckRequredFieldIsFilled(string fieldName, string value, DataRow dr)
		{ 
			if (_requiredFields.Contains(fieldName)
				&& value == string.Empty)
			{
				AddErrorMessage(dr, fieldName, $"В столбце {fieldName} не задано значение. Этот столбец обязателен к заполнению.");
				return false;
			}

			return true;
		}


		private static bool CheckFieldValueIsValid(string fieldName, string value, DataRow dr)
		{
			if (_validValues.ContainsKey(fieldName))
			{
				var validValues = _validValues[fieldName];

				if (!validValues.Contains(value))
				{
					AddErrorMessage(dr, fieldName, $"В столбце {fieldName} задано недопустимое значение {value}");

					return false;
				}
			}

			return true;
		}


		private static bool CheckTagIsValid(string tag, DataRow dr)
		{
			var tagClass = dr["Tag Class Name"].ToString();

			switch (tagClass)
			{
				case "ADSORBER":
				case "BLOWER":
				case "FILTER":
				case "PUMP":
				case "ADSORBER SET":
				case "PUMP SET":
					return CheckEquipmentTagIsValid(tag, tagClass,dr);

				case "COMPRESSOR":
				case "COMPRESSOR SET":
					return CheckEquipmentTagIsValid(tag, "K", dr);

				case "ELECTRIC MOTOR":
					return CheckMotorTagIsValid(tag, dr);	
					
				case "HEAT EXCHANGER SECTION":
					return CheckEquipmentTagIsValid(tag, "EXCHANGER SECTION", dr, true);

				case "STEAM BOILER PACKAGE":
					return CheckEquipmentTagIsValid(tag, "PK", dr, true);

				case "HEAT EXCHANGER":
					return CheckEquipmentTagIsValid(tag, "EXCHANGER", dr);

				case "LINE":
					return CheckLineTagIsValid(tag, dr);

				default:
					return CheckOtherTagIsValid(tag, dr);
			}
		}


        private static bool CheckTagOrderNumber(string text, DataRow dr, bool isSection)
        {
            var number = text.Substring(2, 2).TrimStart('0');

            bool isValid = true;

            if (!int.TryParse(number, out var num))
            {
                AddErrorMessage(dr, "Tag number", $"Порядковый номер оборудования в столбце Tag number {number} не соответствует числовому формату (01..99)");
                isValid = false;
            }

            var spl = text.Split('.');

            if (spl.Length > 1)
            {
				number = spl[spl.Length - 1];

                if (!int.TryParse(spl[spl.Length - 1], out num))
                {
                    AddErrorMessage(dr, "Tag number", $"Код для обозначения составных частей оборудования в столбце Tag number {number} не соответствует числовому формату (1..99)");
                    return false;
                }
            }
            else if (isSection)
            {
                AddErrorMessage(dr, "Tag number", $"Код для обозначения секции оборудования в столбце Tag number не содержит номера секции");
                isValid = false;
            }

            return isValid;
        }


        private static bool CheckOtherTagIsValid(string tag, DataRow dr)
		{
			bool isValid = true;

			var wbs = dr["Wbs Code"].ToString();
			var tagArr = tag.Split('-');

			var title = tagArr[0] + tagArr[2].Substring(0, 2);

            if (tagArr[0] != wbs && tagArr.Length > 2 && title != wbs)
			{
				AddErrorMessage(dr, "Wbs Code", $"Номер титула в столбце Wbs Code {wbs} не соответствует номеру титула, указанному в номере тега ({title})");
				isValid = false;
			}

			if (tagArr.Any(i => i.StartsWith(" ") || i.EndsWith(" ")))
			{
				AddErrorMessage(dr, "Tag number", $"Номер титула содержит пробелы");
				isValid = false;
			}

			return isValid;
		}


		private static bool CheckMotorTagIsValid(string tag, DataRow dr)
		{
			bool isValid = true;

			var wbs = dr["Wbs Code"].ToString();
			var tagArr = tag.Split('-');

			if (tagArr.Length != 3)
			{
				AddErrorMessage(dr, "Tag number", $"В столбце Tag number значение указано в неверном формате");
				return false;
			}

			if (tagArr[2].Length < 4)
			{
				AddErrorMessage(dr, "Tag number", $"В столбце Tag number значение {tagArr[2]} указано в неверном формате");
				return false;
			}

			if (tagArr.Any(i => i.StartsWith(" ") || i.EndsWith(" ")))
			{
				AddErrorMessage(dr, "Tag number", $"Номер титула содержит пробелы");
				isValid = false;
			}

			var title = tagArr[0] + tagArr[2].Substring(0, 2);

			if (title != wbs)
			{
				AddErrorMessage(dr, "Wbs Code", $"Номер титула в столбце Wbs Code {wbs} не соответствует номеру титула, указанному в номере тега ({title})");
				isValid = false;
			}

			if (!_equipmentTypes.Any(i => tagArr[1].StartsWith(i)))
			{
				AddErrorMessage(dr, "Tag Class Name", $"Класс оборудования, указанный в номере тега {tagArr[1]} не соответствует допустимому");
				isValid = false;
			}

			if (!_driveTypes.Any(i => tagArr[1].EndsWith(i)))
			{
				AddErrorMessage(dr, "Tag Class Name", $"Тип привода, указанный в номере тега {tagArr[1]} не соответствует допустимому");
				isValid = false;
			}

			if (!CheckTagOrderNumber(tagArr[2], dr, false))
				isValid = false; 

			return isValid;
		}


		private static bool CheckLineTagIsValid(string tag, DataRow dr)
		{
			bool isValid = true;

			var wbs = dr["Wbs Code"].ToString();
			var tagArr = tag.Split('-');

			if (tagArr.Length != 2)
			{
				AddErrorMessage(dr, "Tag number", $"В столбце Tag number значение указано в неверном формате");
				return false;
			}

			if (tagArr[1].Length < 4)
			{
				AddErrorMessage(dr, "Tag number", $"В столбце Tag number значение {tagArr[1]} указано в неверном формате");
				return false;
			}

			if (tagArr.Any(i => i.StartsWith(" ") || i.EndsWith(" ")))
			{
				AddErrorMessage(dr, "Tag number", $"Номер титула содержит пробелы");
				isValid = false;
			}

			if (tagArr[0] != wbs)
			{
				AddErrorMessage(dr, "Wbs Code", $"Номер титула в столбце Wbs Code {wbs} не соответствует номеру титула, указанному в номере тега ({tagArr[0]})");
				isValid = false;
			}

			if (!int.TryParse(tagArr[1].TrimStart('0'), out var num))
			{
				AddErrorMessage(dr, "Tag number", $"Порядковый номер {tagArr[1]} указан в указан в неверном формате, допустимый формат: (0001..9999)");
				isValid = false;
			}

			return isValid;
		}


		private static bool CheckEquipmentTagIsValid(string tag, string tagClass, DataRow dr, bool isSection = false)
		{
			bool isValid = true;

			var wbs = dr["Wbs Code"].ToString();
			var tagArr = tag.Split('-');

			if (tagArr.Length != 3)
			{ 
				AddErrorMessage(dr, "Tag number", $"Значение титула указано в неверном формате"); 
				return false;
			}

			if (tagArr[2].Length < 4)
			{
				AddErrorMessage(dr, "Tag number", $"В титуле значение {tagArr[2]} указано в неверном формате");
				return false;
			}

			if (tagArr.Any(i => i.StartsWith(" ") || i.EndsWith(" ")))
			{
				AddErrorMessage(dr, "Tag number", $"Номер титула содержит пробелы");
				isValid = false;
			}

			var title = tagArr[0] + tagArr[2].Substring(0, 2);

			if (title != wbs)
			{ 
				AddErrorMessage(dr, "Wbs Code", $"Номер титула в столбце Wbs Code {wbs} не соответствует номеру титула, указанному в номере тега ({title})");
				isValid = false;
			}

			if (!_equipmentTypes.Contains(tagArr[1]) || !tagClass.StartsWith(tagArr[1]))
			{
				AddErrorMessage(dr, "Tag Class Name", $"Класс оборудования, указанный в номере тега {tagArr[1]} не соответствует допустимому для класса ({tagClass})");
				isValid = false;
			}

			if (!CheckTagOrderNumber(tagArr[2], dr, isSection))
				isValid = false;

			return isValid;
		}


		private static DataTable ReadFile(string path)
		{
			string conn = path.EndsWith(".xls")
				? @"Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" + path + ";Extended Properties='Excel 8.0;HRD=Yes';"
				: @"Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" + path + ";Extended Properties='Excel 12.0;HDR=Yes';";

			var dt = new DataTable();
			using (var oleDb = new OleDbConnection(conn))
			{
				var oleAdpt = new OleDbDataAdapter("select * from [data$]", oleDb);
				oleAdpt.Fill(dt);
			}
			return dt;
		}


		private static void AddErrorMessage(DataRow dr, string fieldName, string text)
		{
			var message = $"{_allErrors.Count};{dr["Tag number"]};{dr["Tag Class Name"]};{fieldName};{text}";

			_allErrors.Add(message);
		}
	}
}

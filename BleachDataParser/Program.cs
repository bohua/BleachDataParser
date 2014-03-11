using System;
using NPOI.HSSF.UserModel;
using NPOI.SS.UserModel;
using System.IO;
using Newtonsoft.Json;
using System.Reflection;

namespace xlsParser
{
	class Program
	{
		static void Main(string[] args)
		{
			if (args.Length < 1)
			{
				Console.Write("Error: Please give at least one xls file path!");
				return;
			}
			
			ISheet sheet;
			try
			{
				//FileStream fileStream = new FileStream(@"Templates\report-daily.xls", FileMode.Open);
				FileStream fileStream = new FileStream(args[0], FileMode.Open);
				IWorkbook myWorkbook = new HSSFWorkbook(fileStream);

				sheet = myWorkbook.GetSheetAt(2);

				string json = parseXlsToJson(sheet);
				Console.Write(json);
			}
			catch (Exception ex)
			{
				Console.Write("Error:" + ex.Message);
			}
			//Console.ReadLine();
		}

		static private string parseXlsToJson(ISheet sheet)
		{
			DailyReport report = new DailyReport();
			report.hourly_records = new System.Collections.Generic.List<HourlyRecord>();

			//获取报表日期
			ICell dateCell = sheet.GetRow(1).GetCell(0);
			try
			{
				report.report_date = Convert.ToDateTime(dateCell.ToString());
			}
			catch (Exception ex)
			{
				report.report_date = new DateTime();
			}

			//获取日计数
			report.daily_sum_inbound_throughput_1 = injectIntoClass(sheet.GetRow(29).GetCell(4));
			report.daily_sum_inbound_throughput_2 = injectIntoClass(sheet.GetRow(30).GetCell(4));
			report.daily_sum_inbound_throughput_3 = injectIntoClass(sheet.GetRow(31).GetCell(4));
			report.daily_sum_inbound_total = injectIntoClass(sheet.GetRow(32).GetCell(4));

			report.daily_sum_outbound_throughput_1 = injectIntoClass(sheet.GetRow(33).GetCell(4));
			report.daily_sum_outbound_throughput_2 = injectIntoClass(sheet.GetRow(34).GetCell(4));
			report.daily_sum_outbound_throughput_oldcityregion = injectIntoClass(sheet.GetRow(35).GetCell(4));
			report.daily_sum_outbound_throughput_southregion = injectIntoClass(sheet.GetRow(36).GetCell(4));
			report.daily_sum_outbound_throughput_total = injectIntoClass(sheet.GetRow(37).GetCell(4));

			report.daily_sum_consumption_cl_1 = injectIntoClass(sheet.GetRow(30).GetCell(10));
			report.daily_sum_consumption_cl_2 = injectIntoClass(sheet.GetRow(31).GetCell(10));
			report.daily_sum_consumption_cl_total = injectIntoClass(sheet.GetRow(30).GetCell(12));
			report.daily_sum_consumption_cl_avg = injectIntoClass(sheet.GetRow(30).GetCell(13));

			report.daily_sum_consumption_alun_1 = injectIntoClass(sheet.GetRow(32).GetCell(10));
			report.daily_sum_consumption_alun_2 = injectIntoClass(sheet.GetRow(33).GetCell(10));
			report.daily_sum_consumption_alun_3 = injectIntoClass(sheet.GetRow(34).GetCell(10));
			report.daily_sum_consumption_alun_total = injectIntoClass(sheet.GetRow(32).GetCell(12));
			report.daily_sum_consumption_alun_avg = injectIntoClass(sheet.GetRow(32).GetCell(13));

			report.daily_sum_consumption_alkali_1 = injectIntoClass(sheet.GetRow(35).GetCell(10));
			report.daily_sum_consumption_alkali_2 = injectIntoClass(sheet.GetRow(36).GetCell(10));
			report.daily_sum_consumption_alkali_3 = injectIntoClass(sheet.GetRow(37).GetCell(10));
			report.daily_sum_consumption_alkali_total = injectIntoClass(sheet.GetRow(35).GetCell(12));
			report.daily_sum_consumption_alkali_avg = injectIntoClass(sheet.GetRow(35).GetCell(13));

			report.daily_sum_water_qulity_sediment_NTU_1 = injectIntoClass(sheet.GetRow(29).GetCell(19));
			report.daily_sum_water_qulity_sediment_NTU_2 = injectIntoClass(sheet.GetRow(30).GetCell(19));
			report.daily_sum_water_qulity_filtered_NTU = injectIntoClass(sheet.GetRow(31).GetCell(19));
			report.daily_sum_water_qulity_filtered_PH = injectIntoClass(sheet.GetRow(32).GetCell(19));
			report.daily_sum_water_qulity_filtered_leftCl = injectIntoClass(sheet.GetRow(33).GetCell(19));
			report.daily_sum_water_qulity_outbound_NTU = injectIntoClass(sheet.GetRow(34).GetCell(19));
			report.daily_sum_water_qulity_outbound_PH = injectIntoClass(sheet.GetRow(35).GetCell(19));
			report.daily_sum_water_qulity_outbound_leftCl = injectIntoClass(sheet.GetRow(36).GetCell(19));
			report.daily_sum_water_qulity_outbound_MPA = injectIntoClass(sheet.GetRow(37).GetCell(19));

			report.daily_sum_electron_highprofile_within_mark1 = injectIntoClass(sheet.GetRow(30).GetCell(23));
			report.daily_sum_electron_highprofile_within_mark2 = injectIntoClass(sheet.GetRow(30).GetCell(24));
			report.daily_sum_electron_highprofile_within_marktotal = injectIntoClass(sheet.GetRow(30).GetCell(25));

			report.daily_sum_electron_highprofile_without_mark1 = injectIntoClass(sheet.GetRow(31).GetCell(23));
			report.daily_sum_electron_highprofile_without_mark2 = injectIntoClass(sheet.GetRow(31).GetCell(24));
			report.daily_sum_electron_highprofile_without_marktotal = injectIntoClass(sheet.GetRow(31).GetCell(25));

			report.daily_sum_electron_lowprofile_within_mark1 = injectIntoClass(sheet.GetRow(32).GetCell(23));
			report.daily_sum_electron_lowprofile_within_mark2 = injectIntoClass(sheet.GetRow(32).GetCell(24));
			report.daily_sum_electron_lowprofile_within_marktotal = injectIntoClass(sheet.GetRow(32).GetCell(25));

			report.daily_sum_electron_lowprofile_without_mark1 = injectIntoClass(sheet.GetRow(33).GetCell(23));
			report.daily_sum_electron_lowprofile_without_mark2 = injectIntoClass(sheet.GetRow(33).GetCell(24));
			report.daily_sum_electron_lowprofile_without_marktotal = injectIntoClass(sheet.GetRow(33).GetCell(25));

			report.daily_sum_electron_lowprofile_washroom_mark1 = injectIntoClass(sheet.GetRow(34).GetCell(23));
			report.daily_sum_electron_lowprofile_washroom_mark2 = injectIntoClass(sheet.GetRow(34).GetCell(24));
			report.daily_sum_electron_lowprofile_washroom_marktotal = injectIntoClass(sheet.GetRow(34).GetCell(25));

			report.guard_night = injectStringIntoClass(sheet.GetRow(35).GetCell(20));
			report.guard_morning = injectStringIntoClass(sheet.GetRow(36).GetCell(20));
			report.guard_afternoon = injectStringIntoClass(sheet.GetRow(37).GetCell(20));

			//获取小时级计数
			for (int i = 5; i < 29; i++)
			{
				HourlyRecord newEntry = new HourlyRecord();
				IRow row = sheet.GetRow(i);

				newEntry.record_time = injectTimeIntoClass(row.GetCell(0), report.report_date);

				newEntry.inbound_NTU = injectIntoClass(row.GetCell(1));
				newEntry.inbound_PH = injectIntoClass(row.GetCell(2));
				newEntry.inbound_throughput_1 = injectIntoClass(row.GetCell(3));
				newEntry.inbound_throughput_2 = injectIntoClass(row.GetCell(4));
				newEntry.inbound_throughput_3 = injectIntoClass(row.GetCell(5));

				newEntry.sediment_NTU_1 = injectIntoClass(row.GetCell(7));
				newEntry.sediment_NTU_2 = injectIntoClass(row.GetCell(8));

				newEntry.filtered_NTU = injectIntoClass(row.GetCell(9));
				newEntry.filtered_PH = injectIntoClass(row.GetCell(10));
				newEntry.filtered_leftCl = injectIntoClass(row.GetCell(11));

				newEntry.cleared_leftCl = injectIntoClass(row.GetCell(12));

				newEntry.outbound_NTU = injectIntoClass(row.GetCell(13));
				newEntry.outbound_PH = injectIntoClass(row.GetCell(14));
				newEntry.outbound_leftCl = injectIntoClass(row.GetCell(15));
				newEntry.outbound_MPA = injectIntoClass(row.GetCell(16));
				newEntry.outbound_throughput_1 = injectIntoClass(row.GetCell(17));
				newEntry.outbound_throughput_2 = injectIntoClass(row.GetCell(19));
				newEntry.outbound_throughput_oldtownregion = injectIntoClass(row.GetCell(20));
				newEntry.outbound_throughput_southregion = injectIntoClass(row.GetCell(22));
				newEntry.outbound_total = injectIntoClass(row.GetCell(23));

				newEntry.power_consumption = injectIntoClass(row.GetCell(24));
				newEntry.manufactory_consumption = injectIntoClass(row.GetCell(25));

				report.hourly_records.Add(newEntry);
			}

			try
			{
				string json = JsonConvert.SerializeObject(report);
				return json;
			}
			catch (Exception ex)
			{
				return "Error:" + ex.Message;
			}
		}

		static private decimal injectIntoClass(ICell cell)
		{
			if (cell == null)
			{
				return 0;
			}

			try
			{
				return decimal.Round(decimal.Parse(cell.ToString()), 2);
			}
			catch (Exception ex)
			{
				if (cell.CellType == CellType.Formula && cell.CachedFormulaResultType == CellType.Numeric)
				{
					return decimal.Round((decimal)cell.NumericCellValue, 2);
				}

				return 0;
			}
		}

		static private DateTime injectTimeIntoClass(ICell cell, DateTime dt)
		{
			if (cell == null)
			{
				return new DateTime();
			}

			try
			{
				int hour = int.Parse(cell.ToString());
				return dt.Date.Add(new TimeSpan(hour, 0, 0));
			}
			catch (Exception ex)
			{
				return new DateTime();
			}
		}

		static private string injectStringIntoClass(ICell cell)
		{
			if (cell == null)
			{
				return "";
			}

			try
			{
				return cell.ToString();
			}
			catch (Exception ex)
			{
				return "";
			}
		}
	}
}

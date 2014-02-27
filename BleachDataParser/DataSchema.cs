using System;
using System.Collections.Generic;
using System.Text;

namespace xlsParser
{
	class DailyReport
	{
		public DateTime report_date { get; set; }

		public List<HourlyRecord> hourly_records { get; set; }

		//日水量统计
		//进水量
		public decimal daily_sum_inbound_throughput_1 { get; set; }
		public decimal daily_sum_inbound_throughput_2 { get; set; }
		public decimal daily_sum_inbound_throughput_3 { get; set; }
		public decimal daily_sum_inbound_total { get; set; }
		//供水量
		public decimal daily_sum_outbound_throughput_1 { get; set; }
		public decimal daily_sum_outbound_throughput_2 { get; set; }
		public decimal daily_sum_outbound_throughput_oldcityregion { get; set; }
		public decimal daily_sum_outbound_throughput_southregion { get; set; }
		public decimal daily_sum_outbound_throughput_total { get; set; }

		//日加药量统计
		//氯
		public decimal daily_sum_consumption_cl_1 { get; set; }
		public decimal daily_sum_consumption_cl_2 { get; set; }
		public decimal daily_sum_consumption_cl_total { get; set; }
		public decimal daily_sum_consumption_cl_avg { get; set; }
		//矾
		public decimal daily_sum_consumption_alun_1 { get; set; }
		public decimal daily_sum_consumption_alun_2 { get; set; }
		public decimal daily_sum_consumption_alun_3 { get; set; }
		public decimal daily_sum_consumption_alun_total { get; set; }
		public decimal daily_sum_consumption_alun_avg { get; set; }
		//碱
		public decimal daily_sum_consumption_alkali_1 { get; set; }
		public decimal daily_sum_consumption_alkali_2 { get; set; }
		public decimal daily_sum_consumption_alkali_3 { get; set; }
		public decimal daily_sum_consumption_alkali_total { get; set; }
		public decimal daily_sum_consumption_alkali_avg { get; set; }

		//水质合格率
		//沉淀出水
		public decimal daily_sum_water_qulity_sediment_NTU_1 { get; set; }
		public decimal daily_sum_water_qulity_sediment_NTU_2 { get; set; }
		//滤后水
		public decimal daily_sum_water_qulity_filtered_NTU { get; set; }
		public decimal daily_sum_water_qulity_filtered_PH { get; set; }
		public decimal daily_sum_water_qulity_filtered_leftCl { get; set; }
		//出厂水
		public decimal daily_sum_water_qulity_outbound_NTU { get; set; }
		public decimal daily_sum_water_qulity_outbound_PH { get; set; }
		public decimal daily_sum_water_qulity_outbound_leftCl { get; set; }
		public decimal daily_sum_water_qulity_outbound_MPA { get; set; }

		//日电量统计
		//高配
		//总有功
		public decimal daily_sum_electron_highprofile_within_mark1 { get; set; }
		public decimal daily_sum_electron_highprofile_within_mark2 { get; set; }
		public decimal daily_sum_electron_highprofile_within_marktotal { get; set; }
		//总无功
		public decimal daily_sum_electron_highprofile_without_mark1 { get; set; }
		public decimal daily_sum_electron_highprofile_without_mark2 { get; set; }
		public decimal daily_sum_electron_highprofile_without_marktotal { get; set; }
		//低配
		//总有功
		public decimal daily_sum_electron_lowprofile_within_mark1 { get; set; }
		public decimal daily_sum_electron_lowprofile_within_mark2 { get; set; }
		public decimal daily_sum_electron_lowprofile_within_marktotal { get; set; }
		//总无功
		public decimal daily_sum_electron_lowprofile_without_mark1 { get; set; }
		public decimal daily_sum_electron_lowprofile_without_mark2 { get; set; }
		public decimal daily_sum_electron_lowprofile_without_marktotal { get; set; }
		//反冲洗房
		public decimal daily_sum_electron_lowprofile_washroom_mark1 { get; set; }
		public decimal daily_sum_electron_lowprofile_washroom_mark2 { get; set; }
		public decimal daily_sum_electron_lowprofile_washroom_marktotal { get; set; }

		//值班人员
		//夜班
		public string guard_night { get; set; }
		//白班
		public string guard_morning { get; set; }
		//中班
		public string guard_afternoon { get; set; }
	}

	class HourlyRecord
	{
		public DateTime record_time { get; set; }

		//原水
		public decimal inbound_NTU { get; set; }
		public decimal inbound_PH { get; set; }
		public decimal inbound_throughput_1 { get; set; }
		public decimal inbound_throughput_2 { get; set; }
		public decimal inbound_throughput_3 { get; set; }

		//沉淀出水
		public decimal sediment_NTU_1 { get; set; }
		public decimal sediment_NTU_2 { get; set; }

		//滤后水
		public decimal filtered_NTU { get; set; }
		public decimal filtered_PH { get; set; }
		public decimal filtered_leftCl { get; set; }

		//清水池
		public decimal cleared_leftCl { get; set; }

		//出厂水
		public decimal outbound_NTU { get; set; }
		public decimal outbound_PH { get; set; }
		public decimal outbound_leftCl { get; set; }
		public decimal outbound_MPA { get; set; }
		public decimal outbound_throughput_1 { get; set; }
		public decimal outbound_throughput_2 { get; set; }
		public decimal outbound_throughput_oldtownregion { get; set; }
		public decimal outbound_throughput_southregion { get; set; }
		public decimal outbound_total { get; set; }

		//机组总用电量
		public decimal power_consumption { get; set; }

		//配水单耗
		public decimal manufactory_consumption { get; set; }
	}
}

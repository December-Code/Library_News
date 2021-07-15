using System;
using System.IO;
using System.Text;
using System.Windows.Forms;

namespace AMC
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        ///////////////////////////////Select Calendar//////////////////////////////
        private void MonthBox_SelectedIndexChanged(object sender, EventArgs e)
        {
            DayBox.Text = Calendar.EveryMonthDays(Convert.ToInt32(this.yearText.Text), Convert.ToInt32(this.MonthBox.Text)).ToString();
            switch (Calendar.WhatDay(Convert.ToInt32(this.yearText.Text), Convert.ToInt32(this.MonthBox.Text)))
            {
                case 7:
                    WhichWeek.Text = "星期日";
                    break;
                case 6:
                    WhichWeek.Text = "星期六";
                    break;
                case 5:
                    WhichWeek.Text = "星期五";
                    break;
                case 4:
                    WhichWeek.Text = "星期四";
                    break;
                case 3:
                    WhichWeek.Text = "星期三";
                    break;
                case 2:
                    WhichWeek.Text = "星期二";
                    break;
                case 1:
                    WhichWeek.Text = "星期一";
                    break;
            }
        }
        private void FileName_Click(object sender, EventArgs e)
        {
            if (FileName.Text == "輸入資料夾名...")
                FileName.Text = "";
        }

        private void Excuate_Click(object sender, EventArgs e)
        {
            ///Define
            int years = Convert.ToInt32(this.yearText.Text);
            int month = Convert.ToInt32(this.MonthBox.Text);
            int Days = Convert.ToInt32(this.DayBox.Text);
            int ReadDaysSC = Convert.ToInt32(this.ReadDayBoxS.Text), ReadDaysAD = Convert.ToInt32(this.ReadDayBoxA.Text);
            int StartW_Orignial = 0;
            switch (Convert.ToString(this.WhichWeek.Text))
            {
                case "星期六":
                    StartW_Orignial = 7;
                    break;
                case "星期五":
                    StartW_Orignial = 6;
                    break;
                case "星期四":
                    StartW_Orignial = 5;
                    break;
                case "星期三":
                    StartW_Orignial = 4;
                    break;
                case "星期二":
                    StartW_Orignial = 3;
                    break;
                case "星期一":
                    StartW_Orignial = 2;
                    break;
                case "星期日":
                    StartW_Orignial = 1;
                    break;
            }
            string monthString = "";
            switch (Convert.ToInt32(this.MonthBox.Text))
            {
                case 1:
                    monthString = "January";
                    break;
                case 2:
                    monthString = "February";
                    break;
                case 3:
                    monthString = "March";
                    break;
                case 4:
                    monthString = "April";
                    break;
                case 5:
                    monthString = "May";
                    break;
                case 6:
                    monthString = "June";
                    break;
                case 7:
                    monthString = "July";
                    break;
                case 8:
                    monthString = "August";
                    break;
                case 9:
                    monthString = "September";
                    break;
                case 10:
                    monthString = "October";
                    break;
                case 11:
                    monthString = "November";
                    break;
                case 12:
                    monthString = "December";
                    break;
            }

            string fileNm = FileName.Text.ToString();
            string fmt = "00";

            StringBuilder HtmlTxt = new StringBuilder();
            if (monthString != "January")
                HtmlTxt.AppendLine("<p>&nbsp;</p>");
            HtmlTxt.AppendLine("<p><a id = \"" + years + month.ToString(fmt) + "\" ></a></p>");
            HtmlTxt.AppendLine("<p>&nbsp;</p>");
            HtmlTxt.AppendLine("<table width=\"770\" border=\"1\" cellpadding=\"0\" cellspacing=\"5\" class=\"table\" bordercolor=\"#CC6600\"");
            HtmlTxt.AppendLine("style=\"border-style: outset; \" height=\"677\">");
            HtmlTxt.AppendLine("<tr>");
            HtmlTxt.AppendLine("<td height=\"50px\" colspan=\"2\" class=\"label_title\">" + years + "&emsp;" + month + "&emsp;" + monthString + "</td>");
            HtmlTxt.AppendLine("</tr>");
            HtmlTxt.AppendLine("<tr>");
            HtmlTxt.AppendLine("<td width=\"375\" class=\"text\">");
            HtmlTxt.AppendLine("<div class=\"label_title\"><img");
            HtmlTxt.AppendLine("src=\"http://0-www-o.ntust.edu.tw.sierra.lib.ntust.edu.tw/~lib/studio/img/studio.gif" + "\" alt=\"\" width=\"189\"");
            HtmlTxt.AppendLine("height=\"47\" border=\"0\"></div>");
            HtmlTxt.AppendLine("<table width=\"375\" border=\"0\" cellspacing=\"0\" bordercolorlight=\"#FF9900\" class=\"table_calender\">");
            HtmlTxt.AppendLine("<tr>");
            HtmlTxt.AppendLine("<td>");
            HtmlTxt.AppendLine("<table width=\"375\" border=\"3\" cellpadding=\"3\" cellspacing=\"5\" height=\"289\">");
            HtmlTxt.AppendLine("<tr class=\"label\">");
            HtmlTxt.AppendLine("<td width=\"38\" class=\"label_calender_SC\">Sun</td>");
            HtmlTxt.AppendLine("<td width=\"38\" class=\"label_calender_SC\">Mon</td>");
            HtmlTxt.AppendLine("<td width=\"38\" class=\"label_calender_SC\">Tue</td>");
            HtmlTxt.AppendLine("<td width=\"38\" class=\"label_calender_SC\">Wed</td>");
            HtmlTxt.AppendLine("<td width=\"38\" class=\"label_calender_SC\">Thu</td>");
            HtmlTxt.AppendLine("<td width=\"38\" class=\"label_calender_SC\">Fri</td>");
            HtmlTxt.AppendLine("<td width=\"38\" class=\"label_calender_SC\">Sat</td>");
            HtmlTxt.AppendLine("</tr>");
            ///////////////////////////////////////////////////////////////Date//////////////////////////////////////////////////////////////////
            ///
            int StartW = StartW_Orignial;
            HtmlTxt.AppendLine("<tr class=\"text\">");

            for (int Snap = StartW - 1; Snap > 0; Snap--)
            {
                HtmlTxt.AppendLine("<td class=\"border\">&nbsp;</td>");
            }

            for (int i = 1; i <= Days; i++)
            {
                if (StartW < 7)
                {
                    HtmlTxt.AppendLine("<td class=\"border\"><a href =\"");
                    HtmlTxt.AppendLine("http://0-www-o.ntust.edu.tw.sierra.lib.ntust.edu.tw/~lib/studio/" + years + "/basic/" + fileNm + "/" + "SC"+ (years-2000)+ month.ToString(fmt) + i.ToString(fmt) + ".MP3" + "\">" + i + "</a>");
                    HtmlTxt.AppendLine("</td>");
                    StartW += 1;
                }
                else if (StartW == 7)
                {
                    HtmlTxt.AppendLine("<td class=\"border\"><a href =\"");
                    HtmlTxt.AppendLine("http://0-www-o.ntust.edu.tw.sierra.lib.ntust.edu.tw/~lib/studio/" + years + "/basic/" + fileNm + "/" + "SC"+(years-2000) + month.ToString(fmt) + i.ToString(fmt) + ".MP3" + "\">" + i + "</a>");
                    HtmlTxt.AppendLine("</td>");

                    if (i < Days)
                    {
                        HtmlTxt.AppendLine("</tr>");
                        HtmlTxt.AppendLine("<tr class=\"text\">");
                        StartW = 1;
                    }
                }
            }
            if (StartW != 7)
            {
                for (int Snap = StartW - 1; Snap < 7; Snap++)
                {
                    HtmlTxt.AppendLine("<td class=\"border\">&nbsp;</td>");
                }
            }
            else if (Days - (8 - StartW_Orignial) == 27 && StartW == 7)
            {
                HtmlTxt.AppendLine("<td class=\"border\">&nbsp;</td>");
            }

            HtmlTxt.AppendLine("</tr>");
            HtmlTxt.AppendLine("</table>");
            HtmlTxt.AppendLine("</td>");
            HtmlTxt.AppendLine("</tr>");
            HtmlTxt.AppendLine("</table>");

            HtmlTxt.AppendLine("<br>");
            ////////////////////////////////////////////////////////////Reading//////////////////////////////////////////////////////////////////
            HtmlTxt.AppendLine("<div class=\"label_title\"><img");
            HtmlTxt.AppendLine("src=\"http://0-www-o.ntust.edu.tw.sierra.lib.ntust.edu.tw/~lib/studio/img/reading.gif" + "\" alt=\"\" width=\"189\"");
            HtmlTxt.AppendLine("border=\"0\"></div>");
            HtmlTxt.AppendLine("<table width=\"375\" border=\"3\" cellpadding=\"3\" cellspacing=\"5\">");
            HtmlTxt.AppendLine("<tr bordercolorlight=\"#FF9900\" class=\"label\">");
            HtmlTxt.AppendLine("<td width=\"38\" class=\"label_calender_SC\">Sun</td>");
            HtmlTxt.AppendLine("<td width=\"38\" class=\"label_calender_SC\">Mon</td>");
            HtmlTxt.AppendLine("<td width=\"38\" class=\"label_calender_SC\">Tue</td>");
            HtmlTxt.AppendLine("<td width=\"38\" class=\"label_calender_SC\">Wed</td>");
            HtmlTxt.AppendLine("<td width=\"38\" class=\"label_calender_SC\">Thu</td>");
            HtmlTxt.AppendLine("<td width=\"38\" class=\"label_calender_SC\">Fri</td>");
            HtmlTxt.AppendLine("<td width=\"38\" class=\"label_calender_SC\">Sat</td>");
            HtmlTxt.AppendLine("</tr>");

            StartW = 2;
            int LC = 0;
            HtmlTxt.AppendLine("<tr class=\"text\">");

            HtmlTxt.AppendLine("<td class=\"border\">&nbsp;</td>");//第一天空格
            for (int i = 1; i <= ReadDaysSC; i++)
            {
                if (StartW < 7)
                {
                    HtmlTxt.AppendLine("<td class=\"border\"><a href =\"");
                    HtmlTxt.AppendLine("http://0-www-o.ntust.edu.tw.sierra.lib.ntust.edu.tw/~lib/studio/" + years + "/basic/" + fileNm + "/" + "SC"+(years-2000) + month.ToString(fmt) + (i + Days).ToString(fmt) + ".MP3" + "\">" + i + "</a>");
                    HtmlTxt.AppendLine("</td>");
                    StartW += 1;
                }
                else if (StartW == 7)
                {
                    HtmlTxt.AppendLine("<td class=\"border\"><a href =\"");
                    HtmlTxt.AppendLine("http://0-www-o.ntust.edu.tw.sierra.lib.ntust.edu.tw/~lib/studio/" + years + "/basic/" + fileNm + "/" + "SC"+(years-2000) + month.ToString(fmt) + (i + Days).ToString(fmt) + ".MP3" + "\">" + i + "</a>");
                    HtmlTxt.AppendLine("</td>");
                    LC += 1;
                    if (i < ReadDaysSC)
                    {
                        StartW = 2;
                        HtmlTxt.AppendLine("</tr>");
                        HtmlTxt.AppendLine("<tr class=\"text\">");
                        HtmlTxt.AppendLine("<td class=\"border\">&nbsp;</td>");
                    }
                }
            }
            if (StartW != 7)
            {
                for (int Snap = StartW - 1; Snap < 7; Snap++)
                {
                    HtmlTxt.AppendLine("<td class=\"border\">&nbsp;</td>");
                }
                LC += 1;
            }
            for (int i = LC; i < 3; i++)
            {
                HtmlTxt.AppendLine("</tr>");
                HtmlTxt.AppendLine("<tr class=\"text\">");
                for (int j = 0; j < 7; j++)
                    HtmlTxt.AppendLine("<td class=\"border\">&nbsp;</td>");
            }

            HtmlTxt.AppendLine("</tr>");
            HtmlTxt.AppendLine("</table>");
            HtmlTxt.AppendLine("</td>");

            ////////////////////////////////////////////////////////////Advance//////////////////////////////////////////////////////////////////
            HtmlTxt.AppendLine("<td width=\"375\" class=\"text\" height=\"602\">");
            HtmlTxt.AppendLine("<div class=\"label_title\"><img");
            HtmlTxt.AppendLine("src=\"http://0-www-o.ntust.edu.tw.sierra.lib.ntust.edu.tw/~lib/studio/img/advance.gif" + "\" alt=\"\" width=\"189\"");
            HtmlTxt.AppendLine("height=\"46\" border=\"0\"></div>");
            HtmlTxt.AppendLine("<table width=\"375\" border=\"0\" cellspacing=\"0\" bordercolorlight=\"#FF9900\" class=\"table_calender\">");
            HtmlTxt.AppendLine("<tr>");
            HtmlTxt.AppendLine("<td>");
            HtmlTxt.AppendLine("<table width=\"375\" border=\"3\" cellpadding=\"3\" cellspacing=\"5\" height=\"289\">");
            HtmlTxt.AppendLine("<tr class=\"label\">");
            HtmlTxt.AppendLine("<td width=\"38\" class=\"label_calender_AD\">Sun</td>");
            HtmlTxt.AppendLine("<td width=\"38\" class=\"label_calender_AD\">Mon</td>");
            HtmlTxt.AppendLine("<td width=\"38\" class=\"label_calender_AD\">Tue</td>");
            HtmlTxt.AppendLine("<td width=\"38\" class=\"label_calender_AD\">Wed</td>");
            HtmlTxt.AppendLine("<td width=\"38\" class=\"label_calender_AD\">Thu</td>");
            HtmlTxt.AppendLine("<td width=\"38\" class=\"label_calender_AD\">Fri</td>");
            HtmlTxt.AppendLine("<td width=\"38\" class=\"label_calender_AD\">Sat</td>");
            HtmlTxt.AppendLine("</tr>");
            ////////////////////////////////////////////////////////////Date/////////////////////////////////////////////////////////////////////            
            StartW = StartW_Orignial;///initial
            HtmlTxt.AppendLine("<tr class=\"text\">");
            for (int Snap = StartW - 1; Snap > 0; Snap--)
            {
                HtmlTxt.AppendLine("<td class=\"border\">&nbsp;</td>");
            }
            for (int i = 1; i <= Days; i++)
            {
                if (StartW < 7)
                {
                    HtmlTxt.AppendLine("<td class=\"border\"><a href =\"");
                    HtmlTxt.AppendLine("http://0-www-o.ntust.edu.tw.sierra.lib.ntust.edu.tw/~lib/studio/" + years + "/advance/" + fileNm + "/" + "AD"+(years-2000) + month.ToString(fmt) + i.ToString(fmt) + ".MP3" + "\">" + i + "</a>");
                    HtmlTxt.AppendLine("</td>");
                    StartW += 1;
                }
                else if (StartW == 7)
                {
                    HtmlTxt.AppendLine("<td class=\"border\"><a href =\"");
                    HtmlTxt.AppendLine("http://0-www-o.ntust.edu.tw.sierra.lib.ntust.edu.tw/~lib/studio/" + years + "/advance/" + fileNm + "/" + "AD"+(years-2000) + month.ToString(fmt) + i.ToString(fmt) + ".MP3" + "\">" + i + "</a>");
                    HtmlTxt.AppendLine("</td>");
                    if (i < Days)
                    {
                        HtmlTxt.AppendLine("</tr>");
                        HtmlTxt.AppendLine("<tr class=\"text\">");
                        StartW = 1;
                    }
                }
            }
            if (StartW != 7)
            {
                for (int Snap = StartW - 1; Snap < 7; Snap++)
                {
                    HtmlTxt.AppendLine("<td class=\"border\">&nbsp;</td>");
                }
            }
            else if (Days - (8 - StartW_Orignial) == 27 && StartW == 7)
            {
                HtmlTxt.AppendLine("<td class=\"border\">&nbsp;</td>");
            }

            HtmlTxt.AppendLine("</tr>");
            HtmlTxt.AppendLine("</table>");
            HtmlTxt.AppendLine("</td>");
            HtmlTxt.AppendLine("</tr>");
            HtmlTxt.AppendLine("</table>");

            HtmlTxt.AppendLine("<br>");
            ///////////////////////////////////////////////////////Advance Reading/////////////////////////////////////////////////////////////
            HtmlTxt.AppendLine("<div class=\"label_title\"><img");
            HtmlTxt.AppendLine("src=\"http://0-www-o.ntust.edu.tw.sierra.lib.ntust.edu.tw/~lib/studio/img/reading.gif" + "\" alt=\"\" width=\"189\"");
            HtmlTxt.AppendLine("border=\"0\"></div>");
            HtmlTxt.AppendLine("<table width=\"375\" border=\"3\" cellpadding=\"3\" cellspacing=\"5\">");
            HtmlTxt.AppendLine("<tr bordercolorlight=\"#FF9900\" class=\"label\">");
            HtmlTxt.AppendLine("<td width=\"38\" class=\"label_calender_AD\">Sun</td>");
            HtmlTxt.AppendLine("<td width=\"38\" class=\"label_calender_AD\">Mon</td>");
            HtmlTxt.AppendLine("<td width=\"38\" class=\"label_calender_AD\">Tue</td>");
            HtmlTxt.AppendLine("<td width=\"38\" class=\"label_calender_AD\">Wed</td>");
            HtmlTxt.AppendLine("<td width=\"38\" class=\"label_calender_AD\">Thu</td>");
            HtmlTxt.AppendLine("<td width=\"38\" class=\"label_calender_AD\">Fri</td>");
            HtmlTxt.AppendLine("<td width=\"38\" class=\"label_calender_AD\">Sat</td>");
            HtmlTxt.AppendLine("</tr>");

            StartW = 2;
            LC = 0;
            HtmlTxt.AppendLine("<tr class=\"text\">");
            HtmlTxt.AppendLine("<td class=\"border\">&nbsp;</td>");//第一天空格

            for (int i = 1; i <= ReadDaysAD; i++)
            {
                if (StartW < 7)
                {
                    HtmlTxt.AppendLine("<td class=\"border\"><a href =\"");
                    HtmlTxt.AppendLine("http://0-www-o.ntust.edu.tw.sierra.lib.ntust.edu.tw/~lib/studio/" + years + "/advance/" + fileNm + "/" + "AD"+(years-2000) + month.ToString(fmt) + (i + Days).ToString(fmt) + ".MP3" + "\">" + i + "</a>");
                    HtmlTxt.AppendLine("</td>");
                    StartW += 1;
                    Console.WriteLine(StartW.ToString());
                }
                else if (StartW == 7)
                {
                    HtmlTxt.AppendLine("<td class=\"border\"><a href =\"");
                    HtmlTxt.AppendLine("http://0-www-o.ntust.edu.tw.sierra.lib.ntust.edu.tw/~lib/studio/" + years + "/advance/" + fileNm + "/" + "AD"+(years-2000) + month.ToString(fmt) + (i + Days).ToString(fmt) + ".MP3" + "\">" + i + "</a>");
                    HtmlTxt.AppendLine("</td>");
                    LC += 1;
                    if (i < ReadDaysAD)
                    {
                        StartW = 2;
                        HtmlTxt.AppendLine("</tr>");
                        HtmlTxt.AppendLine("<tr class=\"text\">");
                        HtmlTxt.AppendLine("<td class=\"border\">&nbsp;</td>");
                    }
                }
            }
            if (StartW != 7)
            {
                for (int Snap = StartW - 1; Snap < 7; Snap++)
                {
                    HtmlTxt.AppendLine("<td class=\"border\">&nbsp;</td>");
                }
                LC += 1;
            }
            if (LC < 3)
            {
                for (int i = LC; i < 3; i++)
                {
                    HtmlTxt.AppendLine("</tr>");
                    HtmlTxt.AppendLine("<tr class=\"text\">");
                    for (int j = 0; j < 7; j++)
                        HtmlTxt.AppendLine("<td class=\"border\">&nbsp;</td>");
                }
            }

            //////////////////////////////////////End//////////////////////////////////////////////////
            ///
            HtmlTxt.AppendLine("</tr>");
            HtmlTxt.AppendLine("</table>");
            HtmlTxt.AppendLine("</td>");
            HtmlTxt.AppendLine("</tr>");
            HtmlTxt.AppendLine("</table>");
            HtmlTxt.AppendLine("<p align=\"center\"><a href=\"#top\"><u>Back to top</u></a></p>");


            ////////////////////////////////////Export HTML.txt//////////////////////////////////////
            /////////////////////////////////////////////////////////////////////////////////////////

            StreamWriter ExporText = new StreamWriter("AMC_HTML.txt");
            ExporText.Write(HtmlTxt.ToString());
            if (File.Exists("AMC_HTML.txt"))
                MessageBox.Show("Export successfully!!", "Auto-Generate");
            else
                MessageBox.Show("Error!!", "Auto-Generate");
            ExporText.Close();
        }
    }
    static class Calendar
    {
        private static bool IsLeapYear(int year)//判断某年是不是闰年
        {
            if (year % 4 == 0 && year % 100 != 0 || year % 400 == 0)
            {
                //   Console.WriteLine("shi 闰年");
                return true;
            }
            else
            {
                //   Console.WriteLine("不是 闰年");
                return false;
            }
        }

        public static int WhatDay(int year, int month)//判断从某年某月第一天是星期几
        {
            int num;
            int totalDays = 0;
            for (int i = 1900; i < year; i++)
            {
                if (IsLeapYear(i))
                {
                    totalDays += 366;
                }
                else
                {
                    totalDays += 365;
                }

            }
            for (int j = 1; j < month; j++)
            {
                totalDays += EveryMonthDays(year, j);
            }

            num = totalDays % 7;
            return num + 1;
        }
        public static int EveryMonthDays(int year, int month)//判断某年每个月的天数
        {
            int i = month;
            int monthDay;
            if (i == 1 || i == 3 || i == 5 || i == 7 || i == 8 || i == 10 || i == 12)
            {
                monthDay = 31;
            }

            else if (i == 4 || i == 6 || i == 9 || i == 11)
            {
                monthDay = 30;
            }

            else if (i == 2 && IsLeapYear(year) == true)
            {
                monthDay = 29;
            }
            else
            {
                monthDay = 28;
            }
            return monthDay;
        }
    }
}

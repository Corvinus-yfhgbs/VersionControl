using MnbArfolyam.Entities;
using MnbArfolyam.MnbService;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Windows.Forms.DataVisualization.Charting;
using System.Xml;

namespace MnbArfolyam
{
    public partial class Form1 : Form
    {
        BindingList<RateData> _rateDatas = new BindingList<RateData>();
        BindingList<string> _currs = new BindingList<string>();
        public Form1()
        {
            InitializeComponent();

            GetCurrs();
            
            var start = dateTimePicker1.Value.ToString();
            var end = dateTimePicker2.Value.ToString();
            var value = comboBox1.Text;
            comboBox1.DataSource = _currs;
            RefreshData(start, end, value);

        }
        
        public void GetCurrs()
        {
            var mnbService = new MNBArfolyamServiceSoapClient();

            var request = new GetCurrenciesRequestBody();

            var response = mnbService.GetCurrencies(request);

            var result = response.GetCurrenciesResult;

            XmlDocument xml = new XmlDocument();

            xml.LoadXml(result);

            foreach (XmlElement element in xml.DocumentElement)
            {
                foreach (XmlElement _curr in element.ChildNodes)
                {
                    string curr;
                    curr = _curr.InnerText;
                    _currs.Add(curr);
                }

            }
        }

        public void RefreshData(string start, string end, string value)
        {
            _rateDatas.Clear();
            string xmlString = "";
            using (var sr = new StreamReader("Entities/MnbServiceResult.xml", Encoding.Default))
            {
                while (!sr.EndOfStream)
                {
                    xmlString += sr.ReadLine();
                }
            }
            ReadXML(xmlString);
                //ReadXML(GetExchangeRates(start, end, value));
                dataGridView1.DataSource = _rateDatas;

            chartRateData.DataSource = _rateDatas;
            var series = chartRateData.Series[0];
            series.ChartType = SeriesChartType.Line;
            series.XValueMember = "Date";
            series.YValueMembers = "Value";
            series.BorderWidth = 2;

            var legend = chartRateData.Legends[0];
            legend.Enabled = false;

            var chartArea = chartRateData.ChartAreas[0];
            chartArea.AxisX.MajorGrid.Enabled = false;
            chartArea.AxisY.MajorGrid.Enabled = false;
            chartArea.AxisY.IsStartedFromZero = false;
        }
        public string GetExchangeRates(string start, string end, string value)
        {
            var mnbService = new MNBArfolyamServiceSoapClient();

            var request = new GetExchangeRatesRequestBody()
            {
                currencyNames = value,
                startDate = start,
                endDate = end
            };

            var response = mnbService.GetExchangeRates(request);

            var result = response.GetExchangeRatesResult;

            return result;
        }

        public void ReadXML(string xmlString)
        {
            XmlDocument xml = new XmlDocument();

            xml.LoadXml(xmlString);

            foreach (XmlElement element in xml.DocumentElement)
            {
                var rate = new RateData();
                _rateDatas.Add(rate);

                rate.Date = DateTime.Parse(element.GetAttribute("date"));

                var childElement = (XmlElement)element.ChildNodes[0];
                if (childElement == null)
                    continue;
                rate.Currency = childElement.GetAttribute("curr");

                var unit = decimal.Parse(childElement.GetAttribute("unit"));
                var value = decimal.Parse(childElement.InnerText);
                if (unit != 0) rate.Value = value / unit;
            }
        }

        private void Value_Changed(object sender, EventArgs e)
        {
            var start = dateTimePicker1.Value.ToString();
            var end = dateTimePicker2.Value.ToString();
            var value = comboBox1.SelectedItem.ToString();
            RefreshData(start, end, value);
        }
    }
}

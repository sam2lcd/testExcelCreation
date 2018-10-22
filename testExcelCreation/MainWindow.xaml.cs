using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Navigation;
using System.Windows.Shapes;

/////
//to install in nuget package manager console: Install-Package EPPlus
using OfficeOpenXml;
using OfficeOpenXml.Drawing.Chart;

namespace testExcelCreation
{
    /// <summary>
    /// Interaction logic for MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window
    {
        bool isExcelInstalled;
        FileInfo excelFile = new FileInfo(@"C:\Users\alx\Documents\testExcel_ALX.xlsx");

        //testing listview
        List<string> col1 = new List<string>();
        List<int> col2 = new List<int>();
        List<double> col3 = new List<double>();
        List<string> col4 = new List<string>();

        public MainWindow()
        {
            InitializeComponent();
            isExcelInstalled = Type.GetTypeFromProgID("Excel.Application") != null ? true : false;
        }

        private void Button_Click(object sender, RoutedEventArgs e)
        {
            
            if (isExcelInstalled)
            {
                using (ExcelPackage excel = new ExcelPackage())
                {
                    string workSheetALX = "ALX_testData";
                    excel.Workbook.Worksheets.Add("Worksheet1");
                    excel.Workbook.Worksheets.Add("Worksheet2");
                    excel.Workbook.Worksheets.Add(workSheetALX);

                    // Target a worksheet
                    var ActiveWorksheet = excel.Workbook.Worksheets["Worksheet1"];

                    var headerRow = new List<string[]>()
                    {
                        new string[] { "ID", "First Name", "Last Name", "DOB" }
                    };
                   
                    // write header row data starting from cell A1 (row=1,col=1)
                    ActiveWorksheet.Cells[1,1].LoadFromArrays(headerRow);
                    //header style, cell(startRow,startCol,endRow,endCol) <- select range...
                    ActiveWorksheet.Cells[1,1,1,headerRow[0].Length].Style.Font.Bold = true;
                    ActiveWorksheet.Cells[1, 1, 1, headerRow[0].Length].Style.Font.Size = 14;
                    ActiveWorksheet.Cells[1, 1, 1, headerRow[0].Length].Style.Font.Color.SetColor(System.Drawing.Color.Blue); //add reference for System.Drawing

                    //write many data at once
                    var someRows = new List<object[]>()
                    {
                        new object[] { 1, "ALX", "S", 55.1 },
                        new object[] { -2, "ALX2", "S2", -155.02 }
                    };
                  
                    //write someRows
                    //cells[row,col]
                    ActiveWorksheet.Cells[2, 1].LoadFromArrays(someRows);

                    ///////////////////////////////////////////////
                    //but you could also do it all at once:
                    //write many data at once
                    var moreRows = new List<object[]>()
                    {
                        new object[] { "ID", "First Name", "Last Name", "DOB" },
                        new object[] { 1, "ALX", "S", 55.1 },
                        new object[] { -2, "ALX2", "S2", -155.02 }
                    };
                    ActiveWorksheet = excel.Workbook.Worksheets[2]; //index NOT zero based - as in Cells
                    ActiveWorksheet.Cells[2,4].LoadFromArrays(moreRows);

                    ///////////////////////////////////////
                    //try a chart
                    ActiveWorksheet = excel.Workbook.Worksheets[3]; //index NOT zero based - as in Cells
                    //write many data at once
                    var chartData = new List<object[]>()
                    {
                        new object[] { "X", "Y" },
                    };
                    //fill data
                    int maxX = 20;
                    for(int x = 0; x < maxX; x++)
                    {
                        chartData.Add(new object[] { x, (x*x) });
                    }
                    ActiveWorksheet.Cells[1, 1].LoadFromArrays(chartData);


                    excel.Workbook.Worksheets.Add("chartDemo");
                    ActiveWorksheet = excel.Workbook.Worksheets.Last();
                    ActiveWorksheet.Cells[1, 1].Value = "last workSheet";

                    // add chart
                    var myChart = (ExcelScatterChart)ActiveWorksheet.Drawings.AddChart("chart", eChartType.XYScatterLines);

                    // Define series for the chart  add Yrange ,Xrange (no header)
                    var series = myChart.Series.Add(excel.Workbook.Worksheets[3].Cells[2,2, maxX,2], excel.Workbook.Worksheets[3].Cells[2, 1, maxX, 1]);
                    myChart.Title.Text = "My Chart";
                    myChart.SetSize(500, 500);

                    // Add to 6th row and to the 6th column
                    myChart.SetPosition(2, 0, 4, 0);

                    ///////////////////////////////////////////////////////////////////
                    //add a datetime chart with multimple series
                    excel.Workbook.Worksheets.Add("DateTimeChart");
                    ActiveWorksheet = excel.Workbook.Worksheets.Last();

                    //header
                    ActiveWorksheet.Cells[1, 1].Value = "DateTime";
                    ActiveWorksheet.Cells[1, 2].Value = "sine";
                    ActiveWorksheet.Cells[1, 3].Value = "cosine";
                    //format cells
                    ActiveWorksheet.Cells["A:A"].Style.Numberformat.Format = "d/mm/yyyy hh:mm:ss"; //same as in excel - custom format for datetime
                    ActiveWorksheet.Cells["B:B"].Style.Numberformat.Format = "0.00"; //two decimal places , same as in excel
                    ActiveWorksheet.Cells["C:C"].Style.Numberformat.Format = "0.00"; //two decimal places
                    //fill data
                    chartData.Clear();
                    int maxT = 360;
                    DateTime dt = DateTime.Now;
                    TimeSpan oneSec = new TimeSpan(0, 0, 1);
                    for (int t = 0; t < maxT; t++)
                    {
                        double radians = t * Math.PI / 180;
                        double sine = Math.Sin(radians);
                        double cosine = Math.Cos(radians);
                        chartData.Add(new object[] { dt , sine , cosine  });
                        dt = dt.Add(oneSec);
                    }
                    ActiveWorksheet.Cells[2, 1].LoadFromArrays(chartData);

                    myChart = (ExcelScatterChart)ActiveWorksheet.Drawings.AddChart("DTchart", eChartType.XYScatterLines);

                    // Define series for the chart  add Yrange ,Xrange (no header)
                    var sineSeries = myChart.Series.Add(ActiveWorksheet.Cells[2, 2, maxT, 2], ActiveWorksheet.Cells[2, 1, maxT, 1]);
                    sineSeries.Header = "Sine";
                    var cosineSeries = myChart.Series.Add(ActiveWorksheet.Cells[2, 3, maxT, 3], ActiveWorksheet.Cells[2, 1, maxT, 1]);
                    cosineSeries.Header = "Cosine";
                    myChart.Title.Text = "DT my Chart";

                    myChart.SetSize(800, 500);

                    myChart.SetPosition(1, 0, 4, 0);

                    //save it
                    excel.SaveAs(excelFile);
                }
                //open the file
                System.Diagnostics.Process.Start(excelFile.ToString());
            }
            else
            {
                //excel not installed??

            }
        }

        int counter = 0;
        private void btnListViewTest_Click(object sender, RoutedEventArgs e)
        {
            //add data to listview -> no data binding...
            testListView.Items.Add( new MyItems("small test String", counter++, 3.33,
            "Column 4 here it is a big one String to test the horizontal bar too, let's see if it long enough and what will happen!!") );

        }

        private void testListView_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {//apparently you can only select one 
            string col1 = "";
            int col2 = 0;
            double col3 = 0.0;
            string col4 = "";
                       
            foreach (MyItems i in e.AddedItems)
            {
                col1 = i.col1;
                col2 = i.col2;
                col3 = i.col3;
                col4 = i.col4;
            }

            MessageBox.Show("you clicked on row " + col2.ToString(), "I can see you", MessageBoxButton.OK);
        }
    }

    public class MyItems
    {
        //nothing else needed but the last 4, but you could add more stuff to make your life easier!
        public MyItems(string c1, int c2, double c3, string c4)
        {
            col1 = c1;
            col2 = c2;
            col3 = c3;
            col4 = c4;
        }

        //the following 4 are the same as in the databinding in the designer xml
        public string col1 { get; set; }

        public int col2 { get; set; }

        public double col3 { get; set; }

        public string col4 { get; set; }
    }
}

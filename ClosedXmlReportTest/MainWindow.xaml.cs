using ClosedXML.Report;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Windows;

namespace ClosedXmlReportTest
{
    /// <summary>
    /// Interaction logic for MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window
    {
        private Random _orderCountRandomizer = new Random();
        private Random _priceRandomizer = new Random();

        public MainWindow()
        {
            InitializeComponent();
        }

        private void Button_Click(object sender, RoutedEventArgs e)
        {
            string outputFile = Environment.GetFolderPath(Environment.SpecialFolder.Desktop) + "/dummy_report.xlsx";
            try
            {
                using (var template = new XLTemplate(@"report_template_sold_items.xlsx"))
                {
                    string fromDateStr = DateTime.Now.AddYears(-2).ToString("yyyy-MM-dd");

                    // template.AddVariable("Restaurant", Pref.GetRestaurantName());
                    // template.AddVariable("Branch", Pref.GetBranchName());
                    // template.AddVariable("Phone", Pref.GetBranchPhone());
                    // template.AddVariable("FromDate", fromDateStr);
                    // template.AddVariable("ToDate", DateTime.Now.ToString("yyyy-MM-dd"));

                    template.AddVariable(new SoldItemReport("Greatest Company on Earth",
                        "Booming Branch",
                        "+1-239-234-969",
                        fromDateStr,
                        DateTime.Now.ToString("yyyy-MM-dd"),
                        GetDummyItems()));

                    template.Generate();
                    template.SaveAs(outputFile);
                }
                Console.WriteLine("Path: " + outputFile);
                //Show report
                Process.Start(new ProcessStartInfo(outputFile) { UseShellExecute = true });
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.Message);
                Console.WriteLine(ex.StackTrace);
                Console.WriteLine("Sorry! Cannot generate XLSX now.");
            }
        }

        private List<SoldItem> GetDummyItems()
        {
            List<SoldItem> list = new List<SoldItem>();
            for (int i = 1; i < 51; i++)
                list.Add(new SoldItem(
                    i,
                    "Item " + i,
                    (_orderCountRandomizer.Next() * i) % (100 * i),
                    _priceRandomizer.Next()));
            return list;
        }
    }
}

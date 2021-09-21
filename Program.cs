using System;
using System.Windows;

namespace ExcelData
{
    class Program
    {
        public static void Main(string[] args) {

            new app().Run();
            
        }

        internal class app : Application {
            protected override void OnStartup(StartupEventArgs e) => new WpfApp.MainWindow().Show();
        }
    }
}

using ExcelTemplateLib;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Reflection;
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

namespace ExcelTemplate
{
    public class Class
    {
        public List<Student> Students { get; set; }
        public string ClassName { get; set; }
        public string Teacher { get; set; }
    }

    public class Student
    {
        public string Name { get; set; }
        public int Age { get; set; }
        public string Gender { get; set; }
        public int Score { get; set; }
    }
    /// <summary>
    /// MainWindow.xaml 的交互逻辑
    /// </summary>
    public partial class MainWindow : Window
    {
        public MainWindow()
        {
            InitializeComponent();
            
        }

        private void button_Click(object sender, RoutedEventArgs e)
        {
            var myClass = new Class();
            myClass.ClassName = "Class Four";
            myClass.Teacher = "Robert.R";
            myClass.Students = new List<Student>();
            myClass.Students.Add(new Student()
            {
                Name="Sharon",
                Age=12,
                Gender="Female",
                Score = 70
            });

            myClass.Students.Add(new Student()
            {
                Name = "Robert",
                Age = 13,
                Gender = "Male",
                Score = 65
            });

            var installExecuteFile = Assembly.GetExecutingAssembly().Location;
            var fileInfo = new FileInfo(installExecuteFile);
            var installPath = fileInfo.Directory.FullName;

            var stream = ExcelTemplateHelper.HandleExcel(installPath+@"\ExcelSimple\Simple.xlsx", myClass);
            using (var fileStream = File.Create(installPath + @"\ExcelSimple\1.xlsx"))
            {
                stream.Seek(0, SeekOrigin.Begin);
                stream.CopyTo(fileStream);
            }
        }
    }
}

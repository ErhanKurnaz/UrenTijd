using System.Windows;
using UrenTijd.ui;

namespace UrenTijd
{

    /// <summary>
    /// Interaction logic for MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window
    {
        public struct TimeWindow
        {
            public System.DateTime? from;
            public System.DateTime? until;
        }

        public struct DayFieldsStruct
        {
            public TimeWindow arriving;
            public TimeWindow leaving;
            public TimeWindow working;
            public string workDescription;
            public string workType;
            public bool worked;
        }

        public MainWindow()
        {
            InitializeComponent();
        }

        private DayFieldsStruct ConvertToDayFields(DayFields dayFields)
        {
            DayFieldsStruct dayFieldsStruct = new DayFieldsStruct();

            dayFieldsStruct.arriving.from = dayFields.GoFrom.SelectedTime;
            dayFieldsStruct.arriving.until = dayFields.GoUntil.SelectedTime;

            dayFieldsStruct.leaving.from = dayFields.BackFrom.SelectedTime;
            dayFieldsStruct.leaving.until = dayFields.BackUntil.SelectedTime;

            dayFieldsStruct.working.from = dayFields.WorkFrom.SelectedTime;
            dayFieldsStruct.working.until = dayFields.WorkUntil.SelectedTime;

            dayFieldsStruct.workDescription = dayFields.WorkDescription.Text;
            dayFieldsStruct.workType = dayFields.WorkType.Text;
            dayFieldsStruct.worked = dayFields.worked.IsEnabled;

            return dayFieldsStruct;
        }

        private void Send(object sender, RoutedEventArgs e)
        {
            DayFieldsStruct[] days = new DayFieldsStruct[7];

            days[0] = ConvertToDayFields(Monday);
            days[1] = ConvertToDayFields(Tuesday);
            days[2] = ConvertToDayFields(Wednesday);
            days[3] = ConvertToDayFields(Thursday);
            days[4] = ConvertToDayFields(Friday);
            days[5] = ConvertToDayFields(Saturday);
            days[6] = ConvertToDayFields(Sunday);
        }
    }
}

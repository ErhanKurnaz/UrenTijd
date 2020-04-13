using System.Windows;
using System.Globalization;
using UrenTijd.ui;
using System.ComponentModel;
using System.Threading;
using System;
using System.IO;

namespace UrenTijd
{

    /// <summary>
    /// Interaction logic for MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window, INotifyPropertyChanged
    {
        private int _workWeek;
        private int _workYear;
        private bool _loading = false;
        private bool _showingError = false;

        public event PropertyChangedEventHandler PropertyChanged;

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
            public bool hadBreak;
        }

        public int WorkWeek
        {
            get
            {
                return _workWeek;
            }
            set
            {
                if (value != _workWeek)
                {
                    _workWeek = value;
                    UpdateWeek();
                    OnPropertyChanged("WorkWeek");
                }
            }
        }

        public int WorkYear
        {
            get
            {
                return _workYear;
            }
            set
            {
                if (value != _workYear)
                {
                    _workYear = value;
                    UpdateWeek();
                    OnPropertyChanged("WorkYear");
                }
            }
        }

        public bool Loading
        {
            get
            {
                return _loading;
            }
            set
            {
                if (value != _loading)
                {
                    _loading = value;
                    OnPropertyChanged("Loading");
                }
            }
        }

        public bool ShowingError
        {
            get
            {
                return _showingError;
            }
            set
            {
                if (value != _showingError)
                {
                    _showingError = value;
                    OnPropertyChanged("ShowingError");
                }
            }
        }

        public MainWindow()
        {
            InitializeComponent();
            this.DataContext = this;
            System.DateTime now = System.DateTime.Now;

            var cal = CultureInfo.CurrentCulture.Calendar;
            WorkWeek = cal.GetWeekOfYear(now, CalendarWeekRule.FirstDay, System.DayOfWeek.Monday);
            WorkYear = cal.GetYear(now);
        }

        private void SendMail(DayFieldsStruct[] days)
        {
            try
            {
                Utils.CreateExcel(days, Utils.FirstDateOfWeekISO8601(_workYear, _workWeek));
            }
            catch (InvalidOperationException)
            {
                ShowingError = true;
            }

            Utils.SendMail();
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
            dayFieldsStruct.hadBreak = dayFields.hadBreak.IsChecked ?? false;

            dayFieldsStruct.workDescription = dayFields.WorkDescription.Text;
            dayFieldsStruct.workType = dayFields.WorkType.Text;

            return dayFieldsStruct;
        }

        private void UpdateWeek()
        {
            if (_workYear != 0 && _workWeek != 0)
            {
                System.DateTime beginOfWeek = Utils.FirstDateOfWeekISO8601(_workYear, _workWeek);
                System.DateTime endOfWeek = beginOfWeek.AddDays(6);

                string text = string.Format("Van {0} tot {1}", beginOfWeek.ToString("dd MMM"), endOfWeek.ToString("dd MMM"));
                workWeekBlock.Text = text;
            }
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

            Thread thread = new Thread(() =>
            {
                this.Loading = true;
                this.SendMail(days);
                this.Loading = false;
            });

            thread.Start();
        }

        private void CloseDialog(object sender, RoutedEventArgs e)
        {
            ShowingError = false;
        }

        private void WindowClosed(object sender, EventArgs e)
        {
            if (Directory.Exists(Environment.GetEnvironmentVariable("Temp") + @"\UrenTijd"))
            {
                Directory.Delete(Environment.GetEnvironmentVariable("Temp") + @"\UrenTijd", true);
            }
        }

        protected void OnPropertyChanged(string propertyName)
        {
            PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(propertyName));
        }
    }
}

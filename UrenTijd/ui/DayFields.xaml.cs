using System;
using System.Collections.Generic;
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

namespace UrenTijd.ui
{
    /// <summary>
    /// Interaction logic for DayFields.xaml
    /// </summary>
    public partial class DayFields : UserControl
    {
        public static readonly DependencyProperty DayNameProperty = DependencyProperty.Register("DayName", typeof(string), typeof(UserControl), new PropertyMetadata(""));
        public string DayName { get { return (string)GetValue(DayNameProperty); } set { SetValue(DayNameProperty, value); } }

        public DayFields()
        {
            DataContext = this;
            InitializeComponent();
        }
    }
}

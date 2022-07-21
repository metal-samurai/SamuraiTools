using System;
using System.Collections.Generic;
using System.ComponentModel;
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
using System.Windows.Shapes;

namespace SamuraiTools.Wpf
{
    /// <summary>
    /// Interaction logic for MessageWindow.xaml
    /// </summary>
    public partial class MessageWindow : Window, INotifyPropertyChanged
    {
        private string displayMessage;
        public string DisplayMessage
        {
            get
            {
                return displayMessage;
            }
            set
            {
                displayMessage = value;

                PropertyChanged?.Invoke(this, new PropertyChangedEventArgs("DisplayMessage"));
            }
        }

        public MessageWindow()
        {
            InitializeComponent();

            this.DataContext = this;
        }

        public MessageWindow(string message) : this()
        {
            DisplayMessage = message;
        }

        public event PropertyChangedEventHandler PropertyChanged;
    }
}

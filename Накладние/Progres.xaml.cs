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

namespace Накладние
{
    /// <summary>
    /// Логика взаимодействия для Progres.xaml
    /// </summary>
    public partial class Progres : Window
    {
        private BackgroundWorker backgroundWorker;
        ViewModel document;

        public Progres()
        {
            InitializeComponent();
            backgroundWorker = ((BackgroundWorker)this.FindResource("backgroundWorker"));
            document = new ViewModel();
            Program();
        }

        public void Program()
        {
            progressBar.Maximum = document.FileArray();
            textBlock.Text = document.TextTextBlock;
            backgroundWorker.WorkerReportsProgress = true;
            backgroundWorker.WorkerSupportsCancellation = true;
            backgroundWorker.DoWork += backgroundWorker_DoWork;
            backgroundWorker.RunWorkerAsync();  
        }

        private void backgroundWorker_DoWork(object sender, DoWorkEventArgs e)
        {
            document.Main(backgroundWorker);
            backgroundWorker.CancelAsync();
            this.Close();
        }

        private void backgroundWorker_ProgressChanged(object sender, ProgressChangedEventArgs e)
        {
            progressBar.Value = e.ProgressPercentage;
        }

        private void Button_Click(object sender, RoutedEventArgs e)
        {
            backgroundWorker.CancelAsync(); 
            this.Close();
        }               
    }
}

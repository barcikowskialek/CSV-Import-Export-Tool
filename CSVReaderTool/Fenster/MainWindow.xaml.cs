using CSVReaderTool.Logik;
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

namespace CSVReaderTool
{
    /// <summary>
    /// Interaktionslogik für MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window
    {
        public MainWindow()
        {
            InitializeComponent();
            DataContext = new DatenBearbeitung();
        }

        private void dataGrid_AutoGeneratingColumn(object sender, DataGridAutoGeneratingColumnEventArgs e)
        {
            if (DataContext is CSVReaderTool.Logik.DatenBearbeitung dataContext)
            {
                var spalte = dataContext.Spalten.FirstOrDefault(s => s.Name == e.PropertyName);
                if (spalte != null)
                {
                    e.Column.Header = spalte;
                    e.Column.HeaderTemplate = (DataTemplate)FindResource("HeaderCheckBoxTemplate");
                }
            }


            if (e.Column is DataGridTextColumn textCol)
            {
                var converter = (HttpColor)FindResource("HttpColor");

                var binding = new Binding(e.PropertyName)
                {
                    Converter = converter
                };

                var style = new Style(typeof(TextBlock));
                style.Setters.Add(new Setter(TextBlock.ForegroundProperty, binding));

                textCol.ElementStyle = style;
            }
        }
    }
}

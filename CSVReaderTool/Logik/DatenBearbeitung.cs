using Microsoft.Win32;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.ComponentModel.Design;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Input;

namespace CSVReaderTool.Logik
{
    internal class DatenBearbeitung : INotifyPropertyChanged
    {
        public event PropertyChangedEventHandler PropertyChanged;

        private void OnPropertyChanged(string name)
        {
            PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(name));
        }

        private string _dateiName = "Keine Datei gewählt";
        public string DateiName
        {
            get => _dateiName;
            set
            {
                _dateiName = value;
                OnPropertyChanged(nameof(DateiName));
            }
        }

        public string DateiPfad { get; private set; }

        public ICommand DateiAuswaehlenCommand { get; }

        public DatenBearbeitung()
        {
            DateiAuswaehlenCommand = new MeinCommand(DateiAuswaehlen);
        }

        private void DateiAuswaehlen()
        {
            var dlg = new OpenFileDialog();
            dlg.Filter = "CSV (*.csv)|*.csv|Alle Dateien (*.*)|*.*";

            if (dlg.ShowDialog() == true)
            {
                DateiPfad = dlg.FileName;
                DateiName = Path.GetFileName(dlg.FileName); // nur Name
            }
        }
    }
}
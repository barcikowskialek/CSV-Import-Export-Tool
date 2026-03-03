using Microsoft.VisualBasic.FileIO;
using Microsoft.Win32;
using System;
using System.ComponentModel;
using System.Data;
using System.IO;
using System.Windows.Input;

namespace CSVReaderTool.Logik
{
    internal class DatenBearbeitung : INotifyPropertyChanged
    {
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

        private bool _enabled = false;
        public bool Enabled
        {
            get => _enabled;
            set
            {
                if (_enabled != value)
                {
                    _enabled = value;
                    OnPropertyChanged(nameof(Enabled));
                }
            }
        }

        private DataView _tableView;
        public DataView TableView
        {
            get => _tableView;
            set
            {
                _tableView = value;
                OnPropertyChanged(nameof(TableView));
            }
        }

        public ICommand DateiAuslesenCommand { get; }

        public ICommand DateiAuswaehlenCommand { get; }

        public ICommand DateiExportCommand { get; }

        public DatenBearbeitung()
        {
            DateiAuswaehlenCommand = new MeinCommand(DateiAuswaehlen);
            DateiAuslesenCommand = new MeinCommand(DateiAuslesen);
            DateiExportCommand = new MeinCommand(DateiExport);
        }

        #region INotifyPropertyChanged

        public event PropertyChangedEventHandler PropertyChanged;

        private void OnPropertyChanged(string name)
        {
            PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(name));
        }

        #endregion INotifyPropertyChanged

        #region DateiAuswahl


        private void DateiAuswaehlen()
        {
            var Dialog = new OpenFileDialog();
            Dialog.Filter = "CSV (*.csv)|*.csv|Alle Dateien (*.*)|*.*";

            if (Dialog.ShowDialog() == true)
            {
                DateiPfad = Dialog.FileName;
                DateiName = Path.GetFileNameWithoutExtension(Dialog.FileName);
            }
        }

        #endregion DateiAuswahl

        #region DateiAuslesen

        private void DateiAuslesen()
        {
            if (string.IsNullOrWhiteSpace(DateiPfad) || !File.Exists(DateiPfad))
            {
                System.Windows.MessageBox.Show("Bitte zuerst eine CSV-Datei auswählen.");
                return;
            }

            try
            {
                var table = new DataTable();

                using (var parser = new TextFieldParser(DateiPfad))
                {
                    parser.TextFieldType = FieldType.Delimited;
                    parser.SetDelimiters(";");
                    parser.HasFieldsEnclosedInQuotes = true;

                    string[] headers = parser.ReadFields();
                    foreach (var header in headers)
                        table.Columns.Add(header);

                    while (!parser.EndOfData)
                    {
                        string[] fields = parser.ReadFields();
                        table.Rows.Add(fields);
                    }
                }

                TableView = table.DefaultView;
                Enabled = true;
            }
            catch (Exception ex)
            {
                System.Windows.MessageBox.Show("Fehler beim Einlesen: " + ex.Message);
            }
        }

        #endregion DateiAuslesen

        #region DateiExportieren

        private void DateiExport()
        {
            string StandartOrtner = Path.GetDirectoryName(DateiPfad);

            var Dialog = new Microsoft.Win32.SaveFileDialog();
            Dialog.Filter = "Excel-Datei (*.xlsx)|*.xlsx";
            Dialog.Title = "Excel-Datei speichern";
            Dialog.FileName = DateiName;

            if (!string.IsNullOrWhiteSpace(StandartOrtner))
                Dialog.InitialDirectory = StandartOrtner;

            bool? ok = Dialog.ShowDialog();
            if (ok != true)
                return;

            string exportPath = Dialog.FileName;

            System.Windows.MessageBox.Show("Würde exportieren nach:\n" + exportPath);
        }

        #endregion DateiExportieren
    }
}
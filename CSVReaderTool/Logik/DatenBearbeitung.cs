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
using System.Data;

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
            var dlg = new OpenFileDialog();
            dlg.Filter = "CSV (*.csv)|*.csv|Alle Dateien (*.*)|*.*";

            if (dlg.ShowDialog() == true)
            {
                DateiPfad = dlg.FileName;
                DateiName = Path.GetFileNameWithoutExtension(dlg.FileName);
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
                var lines = File.ReadAllLines(DateiPfad);
                if (lines.Length == 0)
                {
                    System.Windows.MessageBox.Show("Die CSV-Datei ist leer.");
                    return;
                }

                char separator = ';';

                var table = new DataTable();

                var headers = lines[0].Split(separator);
                foreach (var h in headers)
                    table.Columns.Add(h.Trim());

                for (int i = 1; i < lines.Length; i++)
                {
                    if (string.IsNullOrWhiteSpace(lines[i]))
                        continue;

                    var values = lines[i].Split(separator);

                    var row = table.NewRow();
                    for (int c = 0; c < table.Columns.Count; c++)
                    {
                        row[c] = c < values.Length ? values[c].Trim() : "";
                    }
                    table.Rows.Add(row);
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

            var dlg = new Microsoft.Win32.SaveFileDialog();
            dlg.Filter = "Excel-Datei (*.xlsx)|*.xlsx";
            dlg.Title = "Excel-Datei speichern";
            dlg.FileName = DateiName;

            if (!string.IsNullOrWhiteSpace(StandartOrtner))
                dlg.InitialDirectory = StandartOrtner;

            bool? ok = dlg.ShowDialog();
            if (ok != true)
                return;

            string exportPath = dlg.FileName;

            System.Windows.MessageBox.Show("Würde exportieren nach:\n" + exportPath);
        }

        #endregion DateiExportieren
    }
}
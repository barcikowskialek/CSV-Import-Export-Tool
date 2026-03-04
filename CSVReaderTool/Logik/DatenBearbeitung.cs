using CSVReaderTool.Befehle;
using CSVReaderTool.Modelle;
using Microsoft.VisualBasic.FileIO;
using Microsoft.Win32;
using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.ComponentModel;
using System.Data;
using System.IO;
using System.Linq;
using System.Threading;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Input;

namespace CSVReaderTool.Logik
{
    internal class DatenBearbeitung : INotifyPropertyChanged
    {
        private CancellationTokenSource _cancelToken;

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

        private Visibility _loading = Visibility.Hidden;
        public Visibility Loading
        {
            get => _loading;
            set
            {
                if (_loading != value)
                {
                    _loading = value;
                    OnPropertyChanged(nameof(Loading));
                }
            }
        }

        public ObservableCollection<SpaltenAuswahl> Spalten { get; } = new ObservableCollection<SpaltenAuswahl>();

        private DataTable _data;

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
        #region Command
        public ICommand DateiAuslesenCommand { get; }

        public ICommand DateiAuswaehlenCommand { get; }

        public ICommand DateiExportCommand { get; }

        public ICommand AbbrechenCommand { get; }
        public DatenBearbeitung()
        {
            DateiAuswaehlenCommand = new MeinCommand(DateiAuswaehlen);
            DateiAuslesenCommand = new MeinCommand(DateiAuslesen);
            DateiExportCommand = new MeinCommand(DateiExport);
            AbbrechenCommand = new MeinCommand(Abbrechen);
        }
        #endregion Command

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
            OpenFileDialog Dialog = new OpenFileDialog();
            Dialog.Filter = "CSV (*.csv)|*.csv|Alle Dateien (*.*)|*.*";

            if (Dialog.ShowDialog() == true)
            {
                DateiPfad = Dialog.FileName;
                DateiName = Path.GetFileNameWithoutExtension(Dialog.FileName);
            }
        }

        #endregion DateiAuswahl

        #region DateiAuslesen

        private async void DateiAuslesen()
        {
            if (string.IsNullOrWhiteSpace(DateiPfad) || !File.Exists(DateiPfad))
            {
                MessageBox.Show("Bitte zuerst eine CSV-Datei auswählen.");
                return;
            }

            try
            {
                _cancelToken = new CancellationTokenSource();
                CancellationToken token = _cancelToken.Token;
                Loading = Visibility.Visible;

                DataTable table = await Task.Run(() =>
                {
                    DataTable Data = new DataTable();

                    Thread.Sleep(1000);

                    using (TextFieldParser parser = new TextFieldParser(DateiPfad))
                    {
                        parser.TextFieldType = FieldType.Delimited;
                        parser.SetDelimiters(";");
                        parser.HasFieldsEnclosedInQuotes = true;

                        token.ThrowIfCancellationRequested();
                        string[] headers = parser.ReadFields();
                        token.ThrowIfCancellationRequested();

                        if (headers == null) return Data;
                        foreach (string header in headers)
                            Data.Columns.Add(header);

                        while (!parser.EndOfData)
                        {
                            token.ThrowIfCancellationRequested();

                            string[] fields = parser.ReadFields();
                            token.ThrowIfCancellationRequested();

                            if (fields == null) continue;
                            Data.Rows.Add(fields);
                        }
                    }

                    return Data;
                }, token);

                _data = table;

                Spalten.Clear();
                foreach (DataColumn col in table.Columns)
                {
                    Spalten.Add(new SpaltenAuswahl
                    {
                        Name = col.ColumnName,
                        IsChecked = true
                    });
                }

                TableView = table.DefaultView;

                Enabled = true;
            }
            catch (OperationCanceledException)
            {
                
            }
            catch (Exception ex)
            {
                System.Windows.MessageBox.Show("Fehler beim Einlesen: " + ex.Message);

            }
            finally 
            {
                Loading = Visibility.Hidden; 
                _cancelToken?.Dispose(); 
                _cancelToken = null; 
            }
        }

        #endregion DateiAuslesen

        #region DateiExportieren

        private async void DateiExport()
        {
            List<string> SpaltenAuswahl = Spalten.Where(s => s.IsChecked).Select(s => s.Name).ToList();

            if (SpaltenAuswahl.Count == 0)
            {
                MessageBox.Show("Bitte mindestens eine Spalte auswählen.");
                return;
            }

            string exportPath = SpeicherOrtAuswahl();

            if (string.IsNullOrWhiteSpace(exportPath)) 
                return;

            _cancelToken?.Dispose();
            _cancelToken = new CancellationTokenSource();
            await ExcelErstellen(exportPath, SpaltenAuswahl);

            if (!File.Exists(exportPath))
                return;

            _cancelToken?.Dispose();
            _cancelToken = null;

            string argument = "/select, \"" + exportPath + "\"";


            System.Diagnostics.Process.Start("explorer.exe", argument);

        }

        private string SpeicherOrtAuswahl()
        {
            string StandardOrtner = Path.GetDirectoryName(DateiPfad);

            SaveFileDialog Dialog = new SaveFileDialog();
            Dialog.Filter = "Excel-Datei (*.xlsx)|*.xlsx";
            Dialog.Title = "Excel-Datei speichern";
            Dialog.FileName = DateiName;

            if (!string.IsNullOrWhiteSpace(StandardOrtner))
                Dialog.InitialDirectory = StandardOrtner;

            bool? ok = Dialog.ShowDialog();
            if (ok != true)
                return string.Empty;

            string exportPath = Dialog.FileName;
            if (!exportPath.EndsWith(".xlsx", StringComparison.OrdinalIgnoreCase))
                exportPath += ".xlsx";

            return exportPath;
        }

        private async Task ExcelErstellen(string exportPath, List<string> spaltenAuswahl)
        {
            Loading = Visibility.Visible;

            try
            {
                CancellationToken token = _cancelToken.Token;

                await Task.Run(() =>
                {
                    Thread.Sleep(1000);
                    token.ThrowIfCancellationRequested();

                    using (ExcelPackage package = new ExcelPackage())
                    {
                        ExcelWorksheet sheet = package.Workbook.Worksheets.Add(DateiName);

                        for (int c = 0; c < spaltenAuswahl.Count; c++)
                        {
                            token.ThrowIfCancellationRequested();
                            sheet.Cells[1, c + 1].Value = spaltenAuswahl[c];
                        }

                        for (int r = 0; r < _data.Rows.Count; r++)
                        {
                            token.ThrowIfCancellationRequested();

                            for (int c = 0; c < spaltenAuswahl.Count; c++)
                            {
                                token.ThrowIfCancellationRequested();

                                string spaltenName = spaltenAuswahl[c];
                                sheet.Cells[r + 2, c + 1].Value = _data.Rows[r][spaltenName]?.ToString();
                            }
                        }

                        FileInfo datei = new FileInfo(exportPath);
                        package.SaveAs(datei);
                    }
                }, token);
            }
            catch (OperationCanceledException)
            {
            }
            catch (Exception ex)
            {
                MessageBox.Show("Fehler beim Export: " + ex.Message);
            }
            finally
            {
                Loading = Visibility.Hidden;
            }
        }

        #endregion DateiExportieren

        private void Abbrechen()
        {
            _cancelToken?.Cancel();
        }
    }
}
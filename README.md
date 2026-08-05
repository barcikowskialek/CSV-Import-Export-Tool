# CSV Reader Tool

A Windows desktop application for importing CSV files, viewing their contents and exporting selected columns to an Excel file.

The project was developed with C#, WPF and an MVVM-inspired structure. It focuses on file processing, asynchronous operations, data binding and a clear desktop workflow.

## Features

- Select and import CSV files
- Display imported data in a WPF DataGrid
- Automatically generate columns from the CSV headers
- Select individual columns for export
- Export the selected data to an Excel `.xlsx` file
- Perform CSV import and Excel export asynchronously
- Cancel ongoing import and export operations
- Display a loading overlay during longer operations
- Show row numbers in the DataGrid
- Highlight HTTP-related values using a value converter
- Open the exported file location in Windows Explorer
- Handle invalid files and export errors

## Technologies

- C#
- .NET Framework
- WPF
- XAML
- MVVM
- `INotifyPropertyChanged`
- Commands
- `async` / `await`
- `CancellationToken`
- `TextFieldParser`
- EPPlus
- Git

## Project Structure

```text
CSVReaderTool/
├── Ansichten/       # WPF views
├── Befehle/         # Command implementation
├── Converter/       # Value converters
├── Logik/           # Application and data-processing logic
├── Modelle/         # Data models
├── App.xaml
└── CSVReaderTool.csproj

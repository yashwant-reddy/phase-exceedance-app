# PhaseExceedanceApp

A .NET solution for aircraft data filtering, exceedance detection, and reporting, supporting both console and GUI (WinForms/WPF) interfaces.  
Designed for extensibility and code reuse, with clear project separation and future expansion in mind.

---

## 📂 Folder Structure

```plaintext
PhaseExceedanceApp/
│
├── PhaseExceedanceApp.sln                  # Solution file (open this in Visual Studio)
│
├── ExceedanceFilterApp/                    # Console app project folder
│   ├── ExceedanceFilterApp.csproj
│   ├── Program.cs
│   └── ...                                 # Other files for the console app
│
├── PhaseExceedanceFilterApp.Gui/           # WinForms or WPF GUI project folder
│   ├── PhaseExceedanceFilterApp.Gui.csproj
│   ├── Form1.cs / MainWindow.xaml          # Main form/window of GUI
│   └── ...                                 # Other GUI files, resources, etc.
│
├── PhaseFilterApp/                         # (For future use) Another project folder
│   ├── PhaseFilterApp.csproj
│   └── ...
│
└── ExceedanceFilterLib/                    # (Optional, for shared code)
    ├── ExceedanceFilterLib.csproj
    ├── FilterLogic.cs
    └── ...

# Python GUI Application for File Download and Extraction

This project is a Python GUI application that allows users to select directories for Script Location, Workstation Path, and Shared Storage. It includes functionality to download files from a specified URL and extract them, providing a progress bar to indicate the status of the download and extraction process.

## Project Structure

```
python-gui-app
├── src
│   ├── main.py               # Entry point of the application
│   ├── gui
│   │   └── app_gui.py        # GUI layout and functionality
│   ├── utils
│   │   ├── downloader.py      # Functions for downloading and extracting files
│   │   └── paths.py          # Constants for default paths
├── requirements.txt           # Project dependencies
└── README.md                  # Project documentation
```

## Installation

1. Clone the repository:
   ```
   git clone <repository-url>
   cd python-gui-app
   ```

2. Install the required dependencies:
   ```
   pip install -r requirements.txt
   ```

## Usage

1. Run the application:
   ```
   python src/main.py
   ```

2. In the GUI, you will see fields to select:
   - **Script Location**: Default path is set to `\\ad.ccrsb.ca\it-home\IT-SCHOOL-HOME\(UserName)\Login`
   - **Workstation Path**: Default path is set to `C:\CCRCE`
   - **Shared Storage**: Default path is set to `\\ad.ccrsb.ca\xadmin-(SchoolAbbreviation)`

3. After selecting the desired directories, click the **Download** button to initiate the download process. A progress bar will display the status of the download and extraction.

## Contributing

Contributions are welcome! Please open an issue or submit a pull request for any enhancements or bug fixes.

## License

This project is licensed under the MIT License. See the LICENSE file for details.
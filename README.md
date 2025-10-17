# Python Excel ETL GUI

This project provides a graphical user interface (GUI) for performing ETL (Extract, Transform, Load) operations between two Excel files. The application allows users to upload an old Excel file and a new Excel file, visualize the headers, and map the columns from the old file to the new file structure. After mapping, users can transform the data and download the new Excel file with the updated structure.

## Project Structure

```
python-excel-etl-gui
├── src
│   ├── app.py          # Main entry point of the application
│   └── etl.py          # ETL logic for transforming Excel data
├── requirements.txt     # List of dependencies
└── README.md            # Project documentation
```

## Requirements

To run this project, you need to install the following dependencies:

- pandas
- openpyxl
- Tkinter (or PyQt, depending on your choice of GUI library)

You can install the required packages using pip:

```
pip install -r requirements.txt
```

## Usage

1. Clone the repository or download the project files.
2. Install the required dependencies as mentioned above.
3. Run the application by executing the following command:

```
python src/app.py
```

4. Upload the old Excel file and the new Excel file using the provided input elements in the GUI.
5. Visualize the headers of both files and map the columns from the old file to the new file structure.
6. Click the "Transform" button to generate the new Excel file with the selected data and new headings.
7. Download the transformed Excel file.

## Contributing

Contributions are welcome! If you have suggestions for improvements or new features, feel free to open an issue or submit a pull request.

## License

This project is licensed under the MIT License. See the LICENSE file for more details.
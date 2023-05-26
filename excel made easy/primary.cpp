#include <iostream>

using namespace std; 

#include "libxl.h"
#include "FilePath.h"

#include <string>

#include <fstream>

// For read operation: Empty cells are considered invalid
bool isValidCellForRead(libxl::Sheet* sheet, int row, int col) {
    int rowCount = sheet->lastRow();
    int colCount = sheet->lastCol();

    if (row < 1 || row > rowCount || col < 1 || col > colCount)
        return false;

    libxl::CellType cellType = sheet->cellType(row, col);
    return (cellType != libxl::CELLTYPE_EMPTY);
}

// For write operation: Empty cells are considered valid
bool isValidCellForWrite(libxl::Sheet* sheet, int row, int col) {
    int rowCount = sheet->lastRow();
    int colCount = sheet->lastCol();

    if (row < 1 || row > rowCount || col < 1 || col > colCount)
        return false;

    libxl::CellType cellType = sheet->cellType(row, col);
    return true; // All cells are considered valid for writing, including empty cells
}



libxl::Sheet* createNewSheet(libxl::Book* book) {
    string sheetName;
    cout << "Enter the name of the new sheet: ";
    cin.ignore(); // Clear the input buffer
    getline(cin, sheetName);

    // Convert the sheet name to a wide character string
    wstring wSheetName(sheetName.begin(), sheetName.end());

    libxl::Sheet* sheet = book->addSheet(wSheetName.c_str());
    if (!sheet) {
        cout << "Failed to create a new Excel sheet." << endl;
    }
    else {
        cout << "New Excel sheet created." << endl;
    }
    return sheet;
}





libxl::Sheet* loadExistingSheet(libxl::Book* book) {
    cout << "Enter the path of the Excel file to load: ";
    cin >> filePath;

    wstring wideFilePath(filePath.begin(), filePath.end());
    if (book->load(wideFilePath.c_str())) {
        cout << "Excel file loaded successfully." << endl;
        return book->getSheet(0);
    }
    else {
        cout << "Failed to load the Excel file." << endl;
        exit(1);
    }
}


void handleInputError(const  string& errorMessage, libxl::Book* book, const  string& filePath) {
    cerr << "Input error: " << errorMessage << endl;

    // Log the error to a log file
    ofstream logFile("error.log",  ios::app);
    if (logFile.is_open()) {
        logFile << "Input error: " << errorMessage << endl;
        logFile.close();
    }
    else {
        cerr << "Failed to open log file for writing." << endl;
    }

    // Cleanup operations
    if (book) {
        book->release();
         cout << "Excel book released." << endl;
    }

    // Graceful program termination
    exit(1);
}



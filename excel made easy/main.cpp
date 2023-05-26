#include <iostream>
using namespace std; 

#include "libxl.h"
#include <string>

#include "primary.h"
#include "exceli.h"
#include "excelo.h"
#include "FilePath.h"
#include <fstream>


string filePath;


int main() {
    cout << "Choose an option:" << endl;
    cout << "1. Create a new Excel sheet" << endl;
    cout << "2. Load an existing Excel file" << endl;
    int option;
    cin >> option;

    libxl::Book* book;
    libxl::Sheet* sheet;

    if (option == 1) {
        book = xlCreateBook();
        sheet = createNewSheet(book);
        if (!sheet) {
            handleInputError("Failed to create a new sheet.", book, filePath);
            return 1;
        }
    }
    else if (option == 2) {
        book = xlCreateXMLBook();
        sheet = loadExistingSheet(book);
        if (!sheet) {
            handleInputError("Failed to load sheet.", book, filePath);
            return 1;
        }
    }
    else {
        std::cout << "Invalid option selected." << std::endl;
        return 1;
    }

    cout << "Choose an action:" << endl;
    cout << "1. Read data from cells" << endl;
    cout << "2. Write data to cells" << endl;
    int action;
    cin >> action;

    if (action == 1) {
        int startRow, startCol, endRow, endCol;
        cout << "Enter the starting cell (row column): ";
        cin >> startRow >> startCol;
        cout << "Enter the ending cell (row column): ";
        cin >> endRow >> endCol;

        if (isValidCellForRead(sheet, startRow, startCol) && isValidCellForRead(sheet, endRow, endCol)) {
            readDataFromCells(sheet, startRow, startCol, endRow, endCol);
        }
        else {
            cout << "Invalid cell input. Please check the cell coordinates and try again." << endl;
        }

    }
    else if (action == 2) {
        addDataToSheet(sheet);
    }

    else {
        cout << "Invalid action selected." << endl;
    }

    std::wstring wFilePath(filePath.begin(), filePath.end());
    book->save(wFilePath.c_str()); // Save the workbook to the file

    return 0;
}

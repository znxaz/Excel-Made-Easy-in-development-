#include "exceli.h"
#include "excelo.h"
#include <iostream>

using namespace std;

#include "libxl.h"
#include "primary.h"
#include <string>

void addDataToSheet(libxl::Sheet* sheet) {
    int row, col;
    string value;

    cout << "Enter the cell coordinates (row column): ";
    cin >> row >> col;
    cin.ignore(); // Ignore newline character

    wstring wvalue; // Wide character string
    cout << "Enter the value to add: ";
    getline(cin, value);

    wvalue.assign(value.begin(), value.end()); // Convert narrow string to wide string

    if (isValidCellForWrite(sheet, row, col)) {
        sheet->writeStr(row, col, wvalue.c_str()); // Pass the wide character string
        cout << "Value added to cell (" << row << ", " << col << ")." << endl;
    }
    else {
        cout << "Invalid cell input. Please check the cell coordinates and try again." << endl;
    }
}





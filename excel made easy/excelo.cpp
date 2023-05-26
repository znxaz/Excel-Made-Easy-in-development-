#include "excelo.h"

#include <iostream>

using namespace std; 

#include "libxl.h"
#include "primary.h"
#include <string>




void readDataFromCells(libxl::Sheet* sheet, int startRow, int startCol, int endRow, int endCol) {
    for (int row = startRow; row <= endRow; row++) {
        for (int col = startCol; col <= endCol; col++) {
            libxl::CellType cellType = sheet->cellType(row, col);
            if (cellType == libxl::CELLTYPE_NUMBER) {
                double value = sheet->readNum(row, col);
                cout << "Value in cell (" << row << ", " << col << "): " << value << endl;
            }
            else if (cellType == libxl::CELLTYPE_STRING) {
                const wchar_t* value = sheet->readStr(row, col);
                wcout << "String value in cell (" << row << ", " << col << "): " << value << endl;
            }
            else if (cellType == libxl::CELLTYPE_BOOLEAN) {
                bool value = sheet->readBool(row, col);
                cout << "Boolean value in cell (" << row << ", " << col << "): " << (value ? "true" : "false") << endl;
            }
            else if (cellType == libxl::CELLTYPE_BLANK) {
                cout << "Cell (" << row << ", " << col << ") is blank." << endl;
            }
            else if (cellType == libxl::CELLTYPE_ERROR) {
                cout << "Cell (" << row << ", " << col << ") contains an error." << endl;
            }
        }
    }

    cout << "Data reading completed." << endl;
}

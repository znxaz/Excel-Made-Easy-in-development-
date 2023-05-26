#pragma once

#include "libxl.h"


bool isValidCellForRead(libxl::Sheet* sheet, int row, int col); 
bool isValidCellForWrite(libxl::Sheet* sheet, int row, int col); 
libxl::Sheet* createNewSheet(libxl::Book* book);
libxl::Sheet* loadExistingSheet(libxl::Book* book);
void handleInputError(const std::string& errorMessage, libxl::Book* book, const std::string& filePath);
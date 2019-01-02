#ifndef EXCEL_H
#define EXCEL_H

#include <fstream> 
#include <string>
#include <iostream>
#include <streambuf> 

using namespace std;
int excel_rint_col_row(const char *col, const char *row);
float excel_rfloat_col_row(const char *col, const char *row);
char * excel_rstring_col_row(const char *col, const char *row);
void xml_w();
#endif
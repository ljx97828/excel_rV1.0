
#include "Python.h" 
#include <Windows.h>
#include <iostream>
#include <stdio.h>
#include "excel_rw.h"
using namespace std;

int excel_rint_col_row(const char *col, const char *row)
{
	Py_Initialize();//使用python之前，要调用Py_Initialize();这个函数进行初始化
	if (!Py_IsInitialized())
	{
		printf("初始化失败！");
		return 0;
	}
	PyRun_SimpleString("import sys");
	PyRun_SimpleString("sys.path.append('./')");//这一步很重要，修改Python路径
	PyObject * pModule = NULL;//声明变量
	PyObject * pFunc = NULL;// 声明变量
	pModule = PyImport_ImportModule("excel_rw");//这里是要调用的文件名hello.py
	if (pModule == NULL)
	{
		cout << "没找到" << endl;
	}
	pFunc = PyObject_GetAttrString(pModule, "excel_rnum_col_row");//这里是要调用的函数名
	PyObject* args = Py_BuildValue("(ss)", col, row);//给python函数参数赋值
	PyObject* pRet = PyObject_CallObject(pFunc, args);//调用函数
	int res = 0;
	PyArg_Parse(pRet, "i", &res);//转换返回类型
	//cout << "res:" << res << endl;//输出结果
	//Py_Finalize();//调用Py_Finalize，这个根Py_Initialize相对应的。

	return res;
}

float excel_rfloat_col_row(const char *col, const char *row)
{
	Py_Initialize();//使用python之前，要调用Py_Initialize();这个函数进行初始化
	if (!Py_IsInitialized())
	{
		printf("初始化失败！");
		return 0;
	}
	PyRun_SimpleString("import sys");
	PyRun_SimpleString("sys.path.append('./')");//这一步很重要，修改Python路径
	PyObject * pModule = NULL;//声明变量
	PyObject * pFunc = NULL;// 声明变量
	pModule = PyImport_ImportModule("excel_rw");//这里是要调用的文件名hello.py
	if (pModule == NULL)
	{
		cout << "没找到" << endl;
	}
	pFunc = PyObject_GetAttrString(pModule, "excel_rnum_col_row");//这里是要调用的函数名
	PyObject* args = Py_BuildValue("(ss)", col, row);//给python函数参数赋值
	PyObject* pRet = PyObject_CallObject(pFunc, args);//调用函数
	float res = 0;
	PyArg_Parse(pRet, "f", &res);//转换返回类型
								 //cout << "res:" << res << endl;//输出结果
	//Py_Finalize();//调用Py_Finalize，这个根Py_Initialize相对应的。

	return res;
}

char * excel_rstring_col_row(const char *col, const char *row)
{
	Py_Initialize();//使用python之前，要调用Py_Initialize();这个函数进行初始化
	if (!Py_IsInitialized())
	{
		printf("初始化失败！");
		return 0;
	}
	PyRun_SimpleString("import sys");
	PyRun_SimpleString("sys.path.append('./')");//这一步很重要，修改Python路径
	PyObject * pModule = NULL;//声明变量
	PyObject * pFunc = NULL;// 声明变量
	pModule = PyImport_ImportModule("excel_rw");//这里是要调用的文件名hello.py
	if (pModule == NULL)
	{
		cout << "没找到" << endl;
	}
	pFunc = PyObject_GetAttrString(pModule, "excel_rnum_col_row");//这里是要调用的函数名
	PyObject* args = Py_BuildValue("(ss)", col, row);//给python函数参数赋值
	PyObject* pRet = PyObject_CallObject(pFunc, args);//调用函数
	char* res=NULL;
	PyArg_Parse(pRet, "s", &res);//转换返回类型
								 //cout << "res:" << res << endl;//输出结果
								 //Py_Finalize();//调用Py_Finalize，这个根Py_Initialize相对应的
	cout << res;

	return res;
}


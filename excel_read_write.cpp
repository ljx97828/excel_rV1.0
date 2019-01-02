
#include "Python.h" 
#include <Windows.h>
#include <iostream>
#include <stdio.h>
#include "excel_rw.h"
using namespace std;

int excel_rint_col_row(const char *col, const char *row)
{
	Py_Initialize();//ʹ��python֮ǰ��Ҫ����Py_Initialize();����������г�ʼ��
	if (!Py_IsInitialized())
	{
		printf("��ʼ��ʧ�ܣ�");
		return 0;
	}
	PyRun_SimpleString("import sys");
	PyRun_SimpleString("sys.path.append('./')");//��һ������Ҫ���޸�Python·��
	PyObject * pModule = NULL;//��������
	PyObject * pFunc = NULL;// ��������
	pModule = PyImport_ImportModule("excel_rw");//������Ҫ���õ��ļ���hello.py
	if (pModule == NULL)
	{
		cout << "û�ҵ�" << endl;
	}
	pFunc = PyObject_GetAttrString(pModule, "excel_rnum_col_row");//������Ҫ���õĺ�����
	PyObject* args = Py_BuildValue("(ss)", col, row);//��python����������ֵ
	PyObject* pRet = PyObject_CallObject(pFunc, args);//���ú���
	int res = 0;
	PyArg_Parse(pRet, "i", &res);//ת����������
	//cout << "res:" << res << endl;//������
	//Py_Finalize();//����Py_Finalize�������Py_Initialize���Ӧ�ġ�

	return res;
}

float excel_rfloat_col_row(const char *col, const char *row)
{
	Py_Initialize();//ʹ��python֮ǰ��Ҫ����Py_Initialize();����������г�ʼ��
	if (!Py_IsInitialized())
	{
		printf("��ʼ��ʧ�ܣ�");
		return 0;
	}
	PyRun_SimpleString("import sys");
	PyRun_SimpleString("sys.path.append('./')");//��һ������Ҫ���޸�Python·��
	PyObject * pModule = NULL;//��������
	PyObject * pFunc = NULL;// ��������
	pModule = PyImport_ImportModule("excel_rw");//������Ҫ���õ��ļ���hello.py
	if (pModule == NULL)
	{
		cout << "û�ҵ�" << endl;
	}
	pFunc = PyObject_GetAttrString(pModule, "excel_rnum_col_row");//������Ҫ���õĺ�����
	PyObject* args = Py_BuildValue("(ss)", col, row);//��python����������ֵ
	PyObject* pRet = PyObject_CallObject(pFunc, args);//���ú���
	float res = 0;
	PyArg_Parse(pRet, "f", &res);//ת����������
								 //cout << "res:" << res << endl;//������
	//Py_Finalize();//����Py_Finalize�������Py_Initialize���Ӧ�ġ�

	return res;
}

char * excel_rstring_col_row(const char *col, const char *row)
{
	Py_Initialize();//ʹ��python֮ǰ��Ҫ����Py_Initialize();����������г�ʼ��
	if (!Py_IsInitialized())
	{
		printf("��ʼ��ʧ�ܣ�");
		return 0;
	}
	PyRun_SimpleString("import sys");
	PyRun_SimpleString("sys.path.append('./')");//��һ������Ҫ���޸�Python·��
	PyObject * pModule = NULL;//��������
	PyObject * pFunc = NULL;// ��������
	pModule = PyImport_ImportModule("excel_rw");//������Ҫ���õ��ļ���hello.py
	if (pModule == NULL)
	{
		cout << "û�ҵ�" << endl;
	}
	pFunc = PyObject_GetAttrString(pModule, "excel_rnum_col_row");//������Ҫ���õĺ�����
	PyObject* args = Py_BuildValue("(ss)", col, row);//��python����������ֵ
	PyObject* pRet = PyObject_CallObject(pFunc, args);//���ú���
	char* res=NULL;
	PyArg_Parse(pRet, "s", &res);//ת����������
								 //cout << "res:" << res << endl;//������
								 //Py_Finalize();//����Py_Finalize�������Py_Initialize���Ӧ��
	cout << res;

	return res;
}


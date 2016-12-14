#include <QtAxExcelEngine.h>
#include <QtCore/QCoreApplication>


int main(int argc, char *argv[])
{
	QCoreApplication a(argc, argv);

	int ret = 0;

	HANDLE handle_out = ::GetStdHandle(STD_OUTPUT_HANDLE);
	if (handle_out == INVALID_HANDLE_VALUE)
	{
	    return -1;
	}
	//设置屏幕缓冲区和输出屏幕大小
	
	COORD coord = { 121,  301};
	BOOL bret = ::SetConsoleScreenBufferSize(handle_out, coord);
	if (bret == FALSE)
	{
		DWORD ret_error= ::GetLastError();
		//
	}
	SMALL_RECT rect = { 0, 0, 120, 60 };
	bret = ::SetConsoleWindowInfo(handle_out, TRUE, &rect);
	if (bret == FALSE)
	{
	    return -1;
	}

	
	QtAxExcelEngine excel_engine;
	bret = excel_engine.initialize(false);
	if (!bret)
	{
		fprintf(stderr, "initialize excel fail.\n");
	}
	
	//测试相对路径打开，
	bret = excel_engine.open(".\\excel\\example01.xlsx",false);
	if (!bret)
	{
		fprintf(stderr, "Open excel fail.\n");
		return 0;
	}

	//测试使用非预加载 load一个sheet
	fprintf(stderr, "=======================================.\n");
	bret = excel_engine.loadSheet(1, false);
	if (!bret)
	{
		fprintf(stderr, "load excel sheet fail.\n");
	}
	for (int i = 1; i <= excel_engine.rowCount(); ++i)
	{
		for (int j = 1; j <= excel_engine.columnCount(); ++j)
		{
			fprintf(stderr, "cell data row %d column %d data:[%s].\n",
					i,
					j,
					excel_engine.getCell(i, j).toString().toStdString().c_str());
		}
	}

	//测试使用预加载 load一个sheet
	fprintf(stderr, "=======================================.\n");
	bret = excel_engine.loadSheet(1, true);
	if (!bret)
	{
		fprintf(stderr, "load excel sheet fail.\n");
	}
	for (int i = 1; i <= excel_engine.rowCount(); ++i)
	{
		for (int j = 1; j <= excel_engine.columnCount(); ++j)
		{
			fprintf(stderr, "cell data row %d column %d data:[%s].\n", 
					i,
					j,
					excel_engine.getCell(i,j).toString().toStdString().c_str());
		}
	}

	fprintf(stderr, "=======================================.\n");
	
	//这儿故意没有关闭。

	//测试相对路径打开，
	bret = excel_engine.newOne();
	if (!bret)
	{
		fprintf(stderr, "Open excel fail.\n");
	}
	excel_engine.insertSheet("ABCDEFG");

	excel_engine.setCell(1, 1, 1);
	excel_engine.setCell(1, 2, 2);
	excel_engine.setCell(1, 3, 3);
	excel_engine.setCell(1, 4, 4);

	excel_engine.insertSheet("HIJKLMN");
	QVariantList data_list;
	data_list.push_back(QString("A"));
	data_list.push_back(QString("B"));
	data_list.push_back(QString("C"));
	data_list.push_back(QString("D"));
	data_list.push_back(QString("1"));
	data_list.push_back(QString("2"));
	data_list.push_back(QString("3"));
	data_list.push_back(QString("4"));
	data_list.push_back(QString("H"));
	data_list.push_back(QString("I"));
	data_list.push_back(QString("J"));
	data_list.push_back(QString("K"));
	data_list.push_back(QString("5"));
	data_list.push_back(QString("6"));
	data_list.push_back(QString("7"));
	data_list.push_back(QString("8"));
	fprintf(stderr, "=======================================.\n");
	for (int k = 0; k < data_list.size(); ++k)
	{
		fprintf(stderr, " data list %d  string:[%s].\n",
				k,
				data_list[k].toString().toStdString().c_str());
	}
	excel_engine.setRangeCell(1,1,4,4, data_list);

	excel_engine.saveAs(".\\excel\\example02.xlsx");

	excel_engine.finalize();

	return a.quit();
}


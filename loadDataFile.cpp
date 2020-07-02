#include "loadDataFile.h"
#include <windows.h>

LoadDataFile::LoadDataFile(QObject *parent) : QObject(parent)
{
    //_filePath = "";
}

LoadDataFile::~LoadDataFile()
{

}

void LoadDataFile::readData()
{
    int sum;
    //sum = file.len
    //countFile();
    sum = 10;

    for(int i = 0; i < sum; i ++){

    
        Sleep(1000);
        getCount(i, sum);
    }
    emit finish();
}


void LoadDataFile::getCount(int rec, int sum)
{
    emit countStep(rec, sum);
}

void LoadDataFile::doDataOperate()
{

}

#include "loadDataFile.h"
#include <windows.h>

LoadDataFile::LoadDataFile(QObject *parent) : QObject(parent)
{
    //_filePath = "";
}

LoadDataFile::~LoadDataFile()
{

}

void LoadDataFile::readData(QList<MResult*>* results, QString filePath)
{
	_results = results;
	_filePath = filePath;
	_dataOperateType = "readResultData";
}


void LoadDataFile::getCount(int rec, int sum)
{
    emit countStep(rec, sum);
}

void LoadDataFile::doDataOperate()
{
    if(_dataOperateType == "readResultData")
    {

        int sum;
        //sum = file.len
        //countFile();
        sum = 10;

        for(int i = 0; i < sum; i ++)
        {

        
            Sleep(1000);
            getCount(i, sum);
        }
        emit finish();
    }
    else
    {

    }
}

void loadDataFile::readResultsFromFile()
{
	QString appPath = QCoreApplication::applicationDirPath().toStdString().data();
	QFile file(_filePath);
	if (file.exists())
	{
		file.open(QIODevice::ReadOnly);
		QDataStream input(&file);
		QString uid;
		int lenResults;
		input >> _uid >> lenResults;
		for (int ii = 0; ii < lenResults; ii++)
		{
			MResult* rst = _results->at(ii);
			input >> rst->PatternGUID >> rst->_isInitialized >> rst->_validFormCount >> rst->FormResultsLen;
			int len = rst->FormResultsLen;
			for (int i = 0; i < len; i++)
			{
				getCount(i, len);
				MFormResult* mfrst = new MFormResult();
				_results->at(ii)->AddFormResult(mfrst);
				QByteArray mfrstPic;
				input >> mfrst->DeviceIndex >> mfrst->QueueIndex >> mfrst->IsBlankPaper >> mfrst->IsRecognizeSuccess >> mfrst->IsForwardDirection \
					>> mfrst->FormIndex >> mfrst->IsRectified >> mfrst->LenMarkGroupResults >> mfrst->LenImageShotResults >> mfrst->ErrorReason >> mfrst->ImgUrl >> mfrstPic;
				//QPixmap img1(mfrst->ImgUrl);
				QImage img;
				img.loadFromData(mfrstPic);
				QString imgUrl = _path + "/imageGray/" + mfrst->ImgUrl;
				// 避免重名
				QString changedImageFileName;
				QString fileName = imgUrl.left(imgUrl.lastIndexOf("."));
				std::fstream tempFile;
				//ELOGI("start detect  multiple name   " << changedImageFileName.toStdString());
				int imageFileIndex = 0;
				while (true)
				{
					tempFile.open(imgUrl.toLocal8Bit(), std::ios::in);
					if (tempFile)
					{
						//ELOGI("find multiple name " );
						changedImageFileName = fileName + "(" + QString::number(imageFileIndex) + ").jpg";
						imageFileIndex++;
						//ELOGI("find multiple name   " << changedImageFileName.toStdString());
						tempFile.close();
						imgUrl = changedImageFileName;
					}
					else
					{
						//ELOGI("didn't find multiple name ");
						break;
					}
				}
				img.save(imgUrl);
				//查找imgurl中从后往前的第一个/位置
				int first = imgUrl.lastIndexOf("/"); //从后面查找"/"位置
				QString title = imgUrl.right(imgUrl.length() - first - 1); //从右边截取
				mfrst->ImgUrl = title;
				int lenMarks = mfrst->LenMarkGroupResults;
				for (int j = 0; j < lenMarks; j++)
				{
					MGroupResult* mgrt = new MGroupResult();
					input >> mgrt->IsValid >> mgrt->GroupName >> mgrt->TextResult >> mgrt->NumberResult >> mgrt->LenCellResults;
					int lenCells = mgrt->LenCellResults;
					for (int k = 0; k < lenCells; k++)
					{
						MCellResult* mcrst = new MCellResult();
						int kk;
						input >> mcrst->CellIndex >> mcrst->RowIndex >> mcrst->ColIndex >> mcrst->CellName \
							>> mcrst->CellRect.X >> mcrst->CellRect.Y >> mcrst->CellRect.Width >> mcrst->CellRect.Height >> kk;
						mcrst->FillSymbolType = Cm3::FormResult::MarkSymbolResult(kk);
						mgrt->AddCellResult(mcrst);
					}
					mfrst->AddGroupResult(mgrt);
				}
				for (int j = 0; j < mfrst->LenImageShotResults; j++)
				{
					MShotResult* msrst = new MShotResult();
					input >> msrst->GroupName >> msrst->ShotImageByteArray >> msrst->CellRect.X >> msrst->CellRect.Y >> msrst->CellRect.Width >> msrst->CellRect.Height;
					mfrst->AddShotResult(msrst);
				}
				//rst->AddFormResultFromBinaryFile(mfrst);
				//读出标志码识别结果
				MRegionResult* mrgrst = new MRegionResult();
				input >> mrgrst->Name >> mrgrst->X >> mrgrst->Y >> mrgrst->Width >> mrgrst->Height >> mrgrst->Result;
				mfrst->IdentifierResult = mrgrst;
			}
		}
	}
	file.close();
	//MFormResult::saveToFile(Path,*this->FormReuslts);
}



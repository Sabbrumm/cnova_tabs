#include <ActiveQt/qaxobject.h>
#include <ActiveQt/qaxbase.h>

#include <QString>
#include <QFile>
#include <stdexcept>

using namespace std;

#include "ExcelHelper.h"

ExcelExportHelper::ExcelExportHelper(bool closeExcelOnExit)
{
    m_closeExcelOnExit = closeExcelOnExit;
    m_excelApplication = nullptr;
    m_sheet = nullptr;
    m_sheets = nullptr;
    m_workbook = nullptr;
    m_workbooks = nullptr;
    m_excelApplication = nullptr;

    m_excelApplication = new QAxObject( "Excel.Application", 0 );

    if (m_excelApplication == nullptr)
        throw invalid_argument("Failed to initialize interop with Excel (probably Excel is not installed)");

    m_excelApplication->dynamicCall( "SetVisible(bool)", false );
    m_excelApplication->setProperty( "DisplayAlerts", 0);

    m_workbooks = m_excelApplication->querySubObject( "Workbooks" );
    m_workbook = m_workbooks->querySubObject( "Add" );
    m_sheets = m_workbook->querySubObject( "Worksheets" );
    m_sheet = m_sheets->querySubObject( "Add" );
}

void ExcelExportHelper::SetCellValue(int lineIndex, int columnIndex, const QString& value)
{
    QAxObject *cell = m_sheet->querySubObject("Cells(int,int)", lineIndex, columnIndex);
    cell->setProperty("Value",value);
    delete cell;
}

void ExcelExportHelper::SaveAs(const QString& fileName)
{
    if (fileName == "")
        throw invalid_argument("'fileName' is empty!");
    if (fileName.contains("/"))
        throw invalid_argument("'/' character in 'fileName' is not supported by excel!");

    if (QFile::exists(fileName))
    {
        if (!QFile::remove(fileName))
        {
            throw new exception();
        }
    }

    m_workbook->dynamicCall("SaveAs (const QString&)", fileName);
}

ExcelExportHelper::~ExcelExportHelper()
{
    if (m_excelApplication != nullptr)
    {
        if (!m_closeExcelOnExit)
        {
            m_excelApplication->setProperty("DisplayAlerts", 1);
            m_excelApplication->dynamicCall("SetVisible(bool)", true );
        }

        if (m_workbook != nullptr && m_closeExcelOnExit)
        {
            m_workbook->dynamicCall("Close (Boolean)", true);
            m_excelApplication->dynamicCall("Quit (void)");
        }
    }

    delete m_sheet;
    delete m_sheets;
    delete m_workbook;
    delete m_workbooks;
    delete m_excelApplication;
}

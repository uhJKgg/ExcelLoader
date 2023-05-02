#pragma once
#include <windows.h>
#include <string>
#include <iostream>
#include <iomanip>
#include <comutil.h>

#pragma warning (disable:4192)
// Excelを操作するためのタイプライブラリを読みこむ
// Microsoft Office Object Library
#import "libid:2DF8D04C-5BFA-101B-BDE5-00AA0044DE52" no_auto_exclude auto_rename dual_interfaces
// Microsoft Visual Basic for Applications Extensibillity
#import "libid:0002E157-0000-0000-C000-000000000046" dual_interfaces
// Mcrosoft Excel Object Library
#import "libid:00020813-0000-0000-C000-000000000046" no_auto_exclude auto_search auto_rename dual_interfaces


using namespace Excel;

// VBAに必要なCOMオブジェクト管理
class VBAManager
{
public:
    // 初期化
    VBAManager() {
        CoInitialize(NULL);
        excelPtr = Excel::_ApplicationPtr();
        HRESULT hr = excelPtr.CreateInstance(L"Excel.Application");
        // 表示
        excelPtr->Visible[0] = VARIANT_TRUE;
        // 警告無視
        excelPtr->DisplayAlerts[0] = VARIANT_FALSE;
        // ワークブック
        workbooks = excelPtr->Workbooks;
    }
    // 終了
    ~VBAManager() {
        // オブジェクトを参照カウントをデクリメント
        if (worksheetPtr)worksheetPtr.Release();
        if (workbookPtr)workbookPtr.Release();
        if (workbooks)workbooks.Release();
        if (worksheets)worksheets.Release();
        if (excelPtr) {
            excelPtr->Quit();
            excelPtr.Release();
        }
        CoUninitialize();
    }

    // Excel起動設定
    void ExcelStart();
    // Excel起動設定
    void ExcelStart2();
    // ワークシート取得
    Excel::_WorksheetPtr GetWorkSheet(int num)
    {
        if (workbookPtr) {
            worksheets = workbookPtr->Worksheets;
            // ワークシートの取得している
            return worksheets->Item[num];
        }
        return NULL;
    }
private:
    // ブックの追加
    void BookAdd();
    // セルにデータを書き込む
    void CellWrite();
    // ワークブックを保存
    void WorkBookSave(const char* filename);
    // ワークブックを開く
    void WorkBookOpen();
    // セルのデータを読み込む
    void ReadCellData();
private:
    // Excelオブジェクト
    Excel::_ApplicationPtr excelPtr;
    // ワークブック
    Excel::_WorkbookPtr workbookPtr;
    Excel::WorkbooksPtr workbooks;
    // ワークシート
    Excel::_WorksheetPtr worksheetPtr;
    Excel::SheetsPtr worksheets;
};
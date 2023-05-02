#pragma once
#include <windows.h>
#include <string>
#include <iostream>
#include <iomanip>
#include <comutil.h>

#pragma warning (disable:4192)
// Excel�𑀍삷�邽�߂̃^�C�v���C�u������ǂ݂���
// Microsoft Office Object Library
#import "libid:2DF8D04C-5BFA-101B-BDE5-00AA0044DE52" no_auto_exclude auto_rename dual_interfaces
// Microsoft Visual Basic for Applications Extensibillity
#import "libid:0002E157-0000-0000-C000-000000000046" dual_interfaces
// Mcrosoft Excel Object Library
#import "libid:00020813-0000-0000-C000-000000000046" no_auto_exclude auto_search auto_rename dual_interfaces


using namespace Excel;

// VBA�ɕK�v��COM�I�u�W�F�N�g�Ǘ�
class VBAManager
{
public:
    // ������
    VBAManager() {
        CoInitialize(NULL);
        excelPtr = Excel::_ApplicationPtr();
        HRESULT hr = excelPtr.CreateInstance(L"Excel.Application");
        // �\��
        excelPtr->Visible[0] = VARIANT_TRUE;
        // �x������
        excelPtr->DisplayAlerts[0] = VARIANT_FALSE;
        // ���[�N�u�b�N
        workbooks = excelPtr->Workbooks;
    }
    // �I��
    ~VBAManager() {
        // �I�u�W�F�N�g���Q�ƃJ�E���g���f�N�������g
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

    // Excel�N���ݒ�
    void ExcelStart();
    // Excel�N���ݒ�
    void ExcelStart2();
    // ���[�N�V�[�g�擾
    Excel::_WorksheetPtr GetWorkSheet(int num)
    {
        if (workbookPtr) {
            worksheets = workbookPtr->Worksheets;
            // ���[�N�V�[�g�̎擾���Ă���
            return worksheets->Item[num];
        }
        return NULL;
    }
private:
    // �u�b�N�̒ǉ�
    void BookAdd();
    // �Z���Ƀf�[�^����������
    void CellWrite();
    // ���[�N�u�b�N��ۑ�
    void WorkBookSave(const char* filename);
    // ���[�N�u�b�N���J��
    void WorkBookOpen();
    // �Z���̃f�[�^��ǂݍ���
    void ReadCellData();
private:
    // Excel�I�u�W�F�N�g
    Excel::_ApplicationPtr excelPtr;
    // ���[�N�u�b�N
    Excel::_WorkbookPtr workbookPtr;
    Excel::WorkbooksPtr workbooks;
    // ���[�N�V�[�g
    Excel::_WorksheetPtr worksheetPtr;
    Excel::SheetsPtr worksheets;
};
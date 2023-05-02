#include "VBAManager.h"

using namespace Excel;

// �N���ݒ�
void VBAManager::ExcelStart()
{
	if (excelPtr) {
		try {
			// �f�[�^�ǂݍ���
			WorkBookOpen();
			worksheetPtr = GetWorkSheet(1);
			ReadCellData();
			// �ꎞ��~
			::Sleep(10 * 1000);
			// �ŏ���
			excelPtr->WindowState[0] = Excel::xlMinimized;
			std::cout << "�G�N�Z���̋N�����m�F:";
			// ���͑҂�
			std::string s;
			std::getline(std::cin, s);
		}
		catch (_com_error& e) {
			// ��O����
		}
	}
	else {
		std::cerr << "�G�N�Z���N���s��\n";
	}

	std::cout << "�e�X�g�I��\n";
	std::string s;
	std::getline(std::cin, s);
}

void VBAManager::ExcelStart2()
{
	if (excelPtr) {
		try {
			// �f�[�^�ǂݍ���
			BookAdd();
			CellWrite();
			WorkBookSave("Sample2.xlsx");
			// �ꎞ��~
			::Sleep(10 * 1000);
			// �ŏ���
			excelPtr->WindowState[0] = Excel::xlMinimized;
			std::cout << "�G�N�Z���̋N�����m�F:";
			// ���͑҂�
			std::string s;
			std::getline(std::cin, s);
		}
		catch (_com_error& e) {
			// ��O����
		}
	}
	else {
		std::cerr << "�G�N�Z���N���s��\n";
	}

	std::cout << "�e�X�g�I��\n";
	std::string s;
	std::getline(std::cin, s);
}

// �u�b�N�̒ǉ�
void VBAManager::BookAdd()
{
	// �V�KBook�ǉ�
	workbookPtr = workbooks->Add();
}

// �Z���Ƀf�[�^����������
void VBAManager::CellWrite()
{
	// �A�N�e�B�u�ȃV�[�g���擾
	worksheetPtr = workbookPtr->ActiveSheet;

	// ���������
	worksheetPtr->Range["A1"][vtMissing]->Value2 = bstr_t("�G�N�Z�e�X�g");
}

// ���[�N�u�b�N��ۑ�
void VBAManager::WorkBookSave(const char* filename)
{
	bstr_t fileName{ filename };
	workbookPtr->SaveAs(
		filename,
		vtMissing,
		vtMissing,
		vtMissing,
		vtMissing,
		vtMissing,
		Excel::xlNoChange
	);
	// ����
	workbookPtr->Close();
}

// ���[�N�u�b�N���J��
void VBAManager::WorkBookOpen()
{
	workbooks = excelPtr->Workbooks;
	workbookPtr = workbooks->Open("Sample1.xlsx");
}

// �Z���̃f�[�^��ǂݍ���
void VBAManager::ReadCellData()
{
	// 8�s���ǂݍ���ł���
	for (int i = 0; i < 8; ++i) {
		variant_t row{ i + 1 };// �s
		// ���ڂ̂ݓǂݍ���ł���
		variant_t col{ 1 };// ��
		RangePtr cells = worksheetPtr->Cells;
		RangePtr cell = cells->Item[row][col];
		variant_t data = cell->Value2;

		switch (data.vt)
		{
			// ���l
		case VT_R8:
			std::cout << data.dblVal << std::endl;
			break;
			// ������
		case VT_BSTR:
			std::cout << (const char*)bstr_t(data.bstrVal) << '\n';
			break;
			// �f�[�^�Ȃ�
		case VT_EMPTY:
			std::cout << "Empty\n";
			break;
		default:
			break;
		}
	}
}

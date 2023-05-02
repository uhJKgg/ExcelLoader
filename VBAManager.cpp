#include "VBAManager.h"

using namespace Excel;

// 起動設定
void VBAManager::ExcelStart()
{
	if (excelPtr) {
		try {
			// データ読み込み
			WorkBookOpen();
			worksheetPtr = GetWorkSheet(1);
			ReadCellData();
			// 一時停止
			::Sleep(10 * 1000);
			// 最小化
			excelPtr->WindowState[0] = Excel::xlMinimized;
			std::cout << "エクセルの起動を確認:";
			// 入力待ち
			std::string s;
			std::getline(std::cin, s);
		}
		catch (_com_error& e) {
			// 例外処理
		}
	}
	else {
		std::cerr << "エクセル起動不可\n";
	}

	std::cout << "テスト終了\n";
	std::string s;
	std::getline(std::cin, s);
}

void VBAManager::ExcelStart2()
{
	if (excelPtr) {
		try {
			// データ読み込み
			BookAdd();
			CellWrite();
			WorkBookSave("Sample2.xlsx");
			// 一時停止
			::Sleep(10 * 1000);
			// 最小化
			excelPtr->WindowState[0] = Excel::xlMinimized;
			std::cout << "エクセルの起動を確認:";
			// 入力待ち
			std::string s;
			std::getline(std::cin, s);
		}
		catch (_com_error& e) {
			// 例外処理
		}
	}
	else {
		std::cerr << "エクセル起動不可\n";
	}

	std::cout << "テスト終了\n";
	std::string s;
	std::getline(std::cin, s);
}

// ブックの追加
void VBAManager::BookAdd()
{
	// 新規Book追加
	workbookPtr = workbooks->Add();
}

// セルにデータを書き込む
void VBAManager::CellWrite()
{
	// アクティブなシートを取得
	worksheetPtr = workbookPtr->ActiveSheet;

	// 文字列入力
	worksheetPtr->Range["A1"][vtMissing]->Value2 = bstr_t("エクセテスト");
}

// ワークブックを保存
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
	// 閉じる
	workbookPtr->Close();
}

// ワークブックを開く
void VBAManager::WorkBookOpen()
{
	workbooks = excelPtr->Workbooks;
	workbookPtr = workbooks->Open("Sample1.xlsx");
}

// セルのデータを読み込む
void VBAManager::ReadCellData()
{
	// 8行分読み込んでいる
	for (int i = 0; i < 8; ++i) {
		variant_t row{ i + 1 };// 行
		// 一列目のみ読み込んでいる
		variant_t col{ 1 };// 列
		RangePtr cells = worksheetPtr->Cells;
		RangePtr cell = cells->Item[row][col];
		variant_t data = cell->Value2;

		switch (data.vt)
		{
			// 数値
		case VT_R8:
			std::cout << data.dblVal << std::endl;
			break;
			// 文字列
		case VT_BSTR:
			std::cout << (const char*)bstr_t(data.bstrVal) << '\n';
			break;
			// データなし
		case VT_EMPTY:
			std::cout << "Empty\n";
			break;
		default:
			break;
		}
	}
}

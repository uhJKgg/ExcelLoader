#define  WIN32_LEAN_AND_MEAN
#define  _WIN32_WINNT _WIN32_WINNT_WIN7


#include "VBAManager.h"

int main(void)
{

	VBAManager vbaManager;
	int ans = 0;
	std::cout << "エクセルを読み込む：１、エクセルに書き込む：２　＞\n" ;
	std::cin >> ans;
	// 起動
	if(ans == 1)vbaManager.ExcelStart();
	else if(ans == 2)vbaManager.ExcelStart2();

	std::cout << "テストプログラムを終了:";
	std::string s;
	std::getline(std::cin, s);
}
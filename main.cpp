#define  WIN32_LEAN_AND_MEAN
#define  _WIN32_WINNT _WIN32_WINNT_WIN7


#include "VBAManager.h"

int main(void)
{

	VBAManager vbaManager;
	int ans = 0;
	std::cout << "�G�N�Z����ǂݍ��ށF�P�A�G�N�Z���ɏ������ށF�Q�@��\n" ;
	std::cin >> ans;
	// �N��
	if(ans == 1)vbaManager.ExcelStart();
	else if(ans == 2)vbaManager.ExcelStart2();

	std::cout << "�e�X�g�v���O�������I��:";
	std::string s;
	std::getline(std::cin, s);
}
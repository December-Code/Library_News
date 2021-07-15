#include <iostream>
#include <fstream>
#include <iomanip>

using namespace std;
int main()
{
    fstream file;
    int startCount; //從1開始
    int endCount;   // TODO input Endnumber
    cout << "input which number start?" << endl;
    cin >> startCount;
    cout << "input which number End?" << endl;
    cin >> endCount;

    file.open("Web.txt", ios::out);
    if (file.fail())
        cout << "Can't open the file";
    else
    {
        int startCount_add = 0;
        for (int i = endCount; i > startCount - 1; i--)
        {
            if (startCount_add % 5 == 0)
            {
                file << "<tr>" << endl;
            }
            file << "<td><a href=";
            file << "\"https://library-r.ntust.edu.tw/var/file/49/1049/img/2781/Newsletter";
            file << setw(3) << setfill('0') << to_string(i);
            file << ".html\"" << endl;
            file << "target=\"_blank\" title=\"";
            file << setw(3) << setfill('0') << to_string(i);
            file << "\">";
            file << setw(3) << setfill('0') << to_string(i);
            file << "期</a></td>" << endl;
            if (startCount_add % 5 == 4 or i == startCount)
            {
                file << "</tr>" << endl;
            }
            startCount_add += 1;
        }
        file.close();
        cout << "Successly" << endl;
    }
    system("pause");
    return 0;
}
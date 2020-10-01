#include <iostream>
int main()
{
    // commands for making exe

    std::printf("Start making exe....");
    std::system("pyinstaller -F -w -i D:\\YandexDisk\\GitHub\\Python\\todocx\\ico.ico main.py");
    std::printf("\n\nMade it.");
    std::printf("Copy Exapmle and ico");
    std::system("copy ico.ico D:\\YandexDisk\\GitHub\\Python\\todocx\\dist");
    std::system("copy example.docx D:\\YandexDisk\\GitHub\\Python\\todocx\\dist");
    std::printf("\n\n\nFINISH IT!");
    return 0;
}
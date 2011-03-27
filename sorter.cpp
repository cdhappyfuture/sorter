#include "sorter.h"
#include <string.h>
#include <algorithm>        // для std::sort(/**/)

Error::Error()
{}
Error::Error(const char* m)
{
    mess = new char[strlen(m)+1];
    strcpy(mess, m);
}
Error::~Error()
{
    delete[] mess;
}
void Error::text()
{
    return mess;
}

Sorter::Sorter()
{}
Sorter::Sorter(char* filename)
{
    range = NULL;                   // потом можно было проверить что диапазон еще не задан
    rows.reserve(100);              // чтобы быстрее работало, не много памяти съест
    book.Load(filename);
    sheet = book.GetWorkSheet(0);
}

bool Sorter::in_range(Row& row)
{

    //----------------------------------начало задания Ире-------------------------------------
    //-----------------------------------------------------------------------------------------
    // проверяем принадлежность row к range (приватный член класса)
    return 1; //или 0 если не вошел в диапазон
}

void Sorter::fill_it(Range* r)
{
    range = r;
    const int C_CI = 2;                     // С Column Index
    const int D_CI = 3;                     // C, D, K - постоянные
    const int K_CI = 10;
    int E_CI = sheet->GetTotalCols() - 6;   //что-то типа этого, надо тестить
    int F_CI = E_CI + 1;                    // Начиния с E непостоянные, потомучто у
    int GH_CI = E_CI + 2;                   // у цветного файла не фиксированное количество
    int IL_CI = E_CI + 3;                   // столбцов, a подсчеты всегда вставляются в конец
    int MP_CI = E_CI + 4;
    int QT_CI = E_CI + 5;
    Row cur_row;
    // ляляля начинаем цикл по строкам файла
    if (in_range(cur_row))
        rows.push_back(cur_row);
    //конец цикла
    if(rows.size() == 0)
        throw Error("No one row is in range!");
}
    //----------------------------------конец задания Ире-------------------------------------
    //-----------------------------------------------------------------------------------------

    //----------------------------------начало задания Лехе-------------------------------------
    //-----------------------------------------------------------------------------------------
bool comp(Row r1, Row r1)
{
    return r1.E < r2.E;         //например, условие ждем от Ксении, оно может быть очень сложным
}

void Sorter::sort_it()
{
    std::sort(rows.begin(), rows.end(), comp);
}

void save_to_file(char* filename)
{
    // создаем новый xls файл и просто кладем туда содержимое rows
    if(!book.SaveAs(filename))
        throw Error("Error while saving to file");
}
    //----------------------------------конец задания Лехе-------------------------------------
    //-----------------------------------------------------------------------------------------

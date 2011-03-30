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
    // проверяем принадлежность row к range (приватный член класса)
    if ( range->Cb == -1.0 || (row.C >= range->Cb && row.C <= range->Ce) )
        if ( range->Db == -1.0 || (row.D >= range->Db && row.D <= range->De) )
            if ( range->Kb == -1.0 || (row.K >= range->Kb && row.K <= range->Ke) )
                if ( range->Eb == -1.0 || (row.E >= range->Eb && row.E <= range->Ee) )
                    if ( range->Fb == -1.0 || (row.F >= range->Fb && row.F <= range->Fe) )
                        if ( range->GHb == -1.0 || (row.GH >= range->GHb && row.GH <= range->GHe) )
                            if ( range->ILb == -1.0 || (row.IL >= range->ILb && row.IL <= range->ILe) )
                                if ( range->MPb == -1.0 || (row.MP >= range->MPb && row.MP <= range->MPe) )
                                    if ( range->QTb == -1.0 || (row.QT >= range->QTb && row.QT <= range->QTe) )
                                        return 1; // ляпота...
    return 0;
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
    char cur_date[20];
    int row = 1;
    while (++row != sheet->GetTotalRows())
    {
        cell = sheet->Cell(i + 1, date_CI);
        if (cell->GetInteger())
            strcpy(cur_date, DateConvert(cell->GetInteger()));
        else
        {
            strcpy(cur_row.date, cur_date);
            cell = sheet->Cell(row, b_time_CI);

            if (cell->GetDouble() == 0) 
                break; //Если нету даты слева и данных в строке

            strcpy(cur_row.b_time, TimeConvert(cell->GetDouble()));
            cell = sheet->Cell(row, e_time_CI);
            strcpy(cur_row.e_time, TimeConvert(cell->GetDouble()));
            cell = sheet->Cell(row, C_CI);
            cur_row.C = cell->GetDouble();
            cell = sheet->Cell(row, D_CI);
            cur_row.D = cell->GetDouble();
            cell = sheet->Cell(row, K_CI);
            cur_row.K = cell->GetDouble();
            cell = sheet->Cell(row, E_CI);
            cur_row.E = cell->GetDouble();
            cell = sheet->Cell(row, F_CI);
            cur_row.F = cell->GetDouble();
            cell = sheet->Cell(row, GH_CI);
            cur_row.GH = cell->GetDouble();
            cell = sheet->Cell(row, IL_CI);
            cur_row.IL = cell->GetDouble();
            cell = sheet->Cell(row, MP_CI);
            cur_row.MP = cell->GetDouble();
            cell = sheet->Cell(row, QT_CI);
            cur_row.QT = cell->GetDouble();
            // самая важная часть программы =)
            if (in_range(cur_row))
                rows.push_back(cur_row); // Собственно, фильтрация
        }
    }
    if(rows.size() == 0)
        throw Error("No one row is in range!");
}

bool comp(Row r1, Row r2)
{
    return (r1.E < r2.E ? 1 : (r1.F < r2.F ? 1 : (r1.GH < r2.GH ? 1 : 0 ) ) ); //неуверен что сработает
}

void Sorter::sort_it()
{
    std::sort(rows.begin(), rows.end(), comp);
}

void save_to_file(char* filename)
{
    enum {
        date_CI,
        b_time_CI,
        e_time_CI,
        C_CI,
        D_CI,
        K_CI,
        E_CI,
        F_CI,
        GH_CI,
        IL_CI,
        MP_CI,
        QT_CI
    }
    ExcelFormat::BasicExcel new_book;
    new_book.New(1);
    ExcelFormat::BasicExcelWorksheet* new_sheet = new_book.GetWorksheet(0);
    ExcelFormat::BasicExcelCell* new_cell;
    for (int i = 0; i != rows.size(); i++)
    {
        new_cell = new_sheet->Cell(i + 1, date_CI);         // i + 1 = Начинаем со второй строки
        new_cell->Set(rows[i].date);                        // записываем дату в очередную строку
        new_cell = new_sheet->Cell(i + 1, b_time_CI);
        new_cell->Set(rows[i].b_time);
        new_cell = new_sheet->Cell(i + 1, e_time_CI);
        new_cell->Set(rows[i].e_time);
        new_cell = new_sheet->Cell(i + 1, C_CI);
        new_cell->Set(rows[i].C);
        new_cell = new_sheet->Cell(i + 1, D_CI);
        new_cell->Set(rows[i].D);
        new_cell = new_sheet->Cell(i + 1, K_CI);
        new_cell->Set(rows[i].K);
        new_cell = new_sheet->Cell(i + 1, E_CI);
        new_cell->Set(rows[i].E);
        new_cell = new_sheet->Cell(i + 1, F_CI);
        new_cell->Set(rows[i].F);
        new_cell = new_sheet->Cell(i + 1, GH_CI);
        new_cell->Set(rows[i].GH);
        new_cell = new_sheet->Cell(i + 1, IL_CI);
        new_cell->Set(rows[i].IL);
        new_cell = new_sheet->Cell(i + 1, MP_CI);
        new_cell->Set(rows[i].MP);
        new_cell = new_sheet->Cell(i + 1, QT_CI);
        new_cell->Set(rows[i].QT);
    }    
    if(!book.SaveAs(filename))
        throw Error("Error while saving to file");
}

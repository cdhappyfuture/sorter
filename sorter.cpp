#include "sorter.h"
#include <string.h>
#include <cstdio>
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
    //delete[] mess;
}
char* Error::text()
{
    return mess;
}

Sorter::Sorter()
{}
Sorter::Sorter(const char* filename)
{
    range = NULL;                   // потом можно было проверить что диапазон еще не задан
    rows.reserve(100);              // чтобы быстрее работало, не много памяти съест
    book.Load(filename);
    sheet = book.GetWorksheet(0);
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


struct Date 
{
    int y;
    int m;
    int d;
};

Date* DaysToDate(int days)
{//принимает количество дней и возвращает дату в виде строки
    Date* date = new Date;
    date->m = 0;
    const int BEGIN = 1900;
    int dm[][12] = {{31, 29, 31, 30, 31, 30, 31, 31, 30, 31, 30, 31},
        {31, 28, 31, 30, 31, 30, 31, 31, 30, 31, 30, 31}};
    date->y = days*4 / (static_cast<int>(365.25*4));                                         //количество лет в полученом периоде
    int usegod = (date->y % 4 ? 0 : 1);                               //если год вясокосный то 0 иначе 1
    date->d = days - date->y*365 - date->y/4 - usegod;                        //количество дней в неполном году
    while (date->d > dm[usegod][date->m])                               //пока полный месяц
        date->d -= dm[usegod][date->m++];                                 //вычитаем его длину из неполного года
    date->m++;
    date->y += BEGIN;
    return date;
}

char* DaysToStr(int days)
{
    Date* dt = DaysToDate(days);
    char* date_str = new char[11];
    std::sprintf(date_str, "%d.%d.%d", dt->d, dt->m, dt->y);
    delete dt;
    return date_str;
}

char* TimeConvert(double d_time)
{
    char* s_time = new char[6];
    int h = 24 * d_time;
    int m = (d_time - (h / 24.0) + 0.00001) * 24 * 60;
    std::sprintf(s_time, "%d:%d", h, m);
    return s_time;
}

void Sorter::fill_it(Range* r)
{
    range = r;
    // индексы стлобцов в исходном файле
    const int date_CI = 0,
          b_time_CI = 5,
          e_time_CI = 6,
          C_CI = 2,                         // С Column Index
          D_CI = 3,                         // C, D, K ,date, b_time, e_time - постоянные
          K_CI = 10;
    int E_CI = sheet->GetTotalCols() - 7,
        F_CI = E_CI + 1,                    // Начиния с E непостоянные, потомучто у
        GH_CI = E_CI + 2,                   // у цветного файла не фиксированное количество
        IL_CI = E_CI + 3,                   // столбцов, a подсчеты всегда вставляются в конец
        MP_CI = E_CI + 4,
        QT_CI = E_CI + 5;
    Row cur_row;
    char* cur_date = NULL;
    int row = 1;
    while (++row != sheet->GetTotalRows())
    {
        cell = sheet->Cell(row, date_CI);
        if (cell->GetInteger())
        {
            delete cur_date;
            cur_date= DaysToStr(cell->GetInteger());
        }
        
        else
        {
            cur_row.date = new char[strlen(cur_date)+1];
            strcpy(cur_row.date, cur_date);
            cell = sheet->Cell(row, b_time_CI);

            if (cell->GetDouble() == 0) 
                break; //Если нету даты слева и данных в строке

            cur_row.b_time = new char[6];
            cur_row.e_time = new char[6];
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
    delete cur_date;
    if(rows.size() == 0)
        throw Error("No one row is in range!");
}

bool comp( const Row& r1, const Row& r2)
{
    return r1.E < r2.E;
}

void Sorter::sort_it()
{
    std::sort(rows.begin(), rows.end(), comp);
}

void Sorter::save_to_file(const char* filename)
{
    // индексы столбцов в новом файле
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
    };
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
    if(!new_book.SaveAs(filename))
        throw Error("Error while saving to file");
}

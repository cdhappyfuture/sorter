#ifndef SORTER_H
#define SORTER_H
#include <vector>
#include <ExcelFormat.h>

struct Range
{
    int Cb, Ce;
    int Db, De;
    int Kb, Ke;
    int Eb, Ee;
    int Fb, Fe;
    int GHb, GHe;
    int ILb, ILe;
    int MPb, MPe;
    int QTb, QTe;
};

struct Row
{
    int H;//жду ответа от Ксении
};
class Error
{
    char* mess;
    Error();
public:
    Error(char* m);
    ~Error();
    text();
};
class Sorter
{
public:                             //интерфейс класса для программы

    Sorter(char* filename);         // инициализируем объект именем xls файла

    fill_it(Range* range);          // заполняем вектор строк данными из исходного
                                    //файла, входящими в диапазоны range

    sort_it(int fields);            //Сортируем выборку по полям указанным в
                                    //fields (попробуем заюзать битовые операции)

    save_to_file(char* filename);   //Сохраняем отсортированный результат в filename

private:                            //остальное
    ExcelFormat::BasicExcel book;
    ExcelFormat::BasicExcelWorksheet* sheet;
    ExcelFormat::BasicExcelCell* cell;
    std::vector<Row> row;
    Sorter();                       // В private чтобы нельзя было создать объект
                                    //без инициализации xls файлом
};
/* Ренату вставить в гуи примерной такой текст
#include <sorter.h>
// blablabla
// ...
try
{
// должен быть выбран файл исходник
Sorter sorter("файл исходник"); //char*, если уж стринг то s.c_str()
//Заполняешь структуру "Range range;" данными из спинбоксов
sorter.fill_it(&range);
sorter.sort_it();
// выбирается файл куда сохранить результат
sorter.save_to_file();
//showmessage кутишный тип все получилось
}
catch (Error err)
{
//showmessage кутишный с текстом err.text()
}
*/
#endif

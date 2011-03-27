#ifndef SORTER_H
#define SORTER_H
#include <vector>
#include <ExcelFormat.h>


struct Row
{                       //Структура "строка" нужных нам данных из цветного файла
    // те что были
    char*  date;        //  дата стоит перед блоками записей, но нужно при заполнении
                        //  вектора строк в каждую запись добавлять дату
    char*  b_time;      // начало диапазона
    char*  e_time;      // конец диапазона
    double  C;          // обводненность % ваще оно целое, но для удобства сделал все дабл
    double  D;          // Расход жидкости (режимная карта), [м3/сут]
    double  K;          // Расход газа     (RVG), [м3/сут]
    // 6 добавленных средних
    double  E;          // Температура
    double  F;          //  давление
    double  GH;         // обводненность OW
    double  IL;         // частота Доплера
    double  MP;         // газонасыщенность
    double  QT;         // загадочный N
};

struct Range            // Диапазон значений для выборки из цветного файла
{                       // Заполняется из GUI, передается методу Sorter::fill_it(Range* range)
    double  Cb, Ce;
    double  Db, De;
    double  Kb, Ke;
    double  Eb, Ee;
    double  Fb, Fe;
    double  GHb, GHe;
    double  ILb, ILe;
    double  MPb, MPe;
    double  QTb, QTe;
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
#endif

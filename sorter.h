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
    Error(const char* m);
    ~Error();
    char* text();
};

class Sorter
{
public:                             //интерфейс класса для программы

    Sorter(char* filename);         // инициализируем объект именем xls файла

    void fill_it(Range* range);          // заполняем вектор строк данными из исходного
                                        //файла, входящими в диапазоны range

    void sort_it();            //Сортируем выборку по полям указанным в
                                        //fields (ждем ответа от Ксении)

    void save_to_file(char* filename);   //Сохраняем отсортированный результат в filename

private:                            //остальное
    Range* range;
    ExcelFormat::BasicExcel book;
    ExcelFormat::BasicExcelWorksheet* sheet;
    ExcelFormat::BasicExcelCell* cell;
    std::vector<Row> rows;
    Sorter();                       // В private чтобы нельзя было создать объект
    bool in_range(Row& row);             // Проверка на принадлежности строки диапазону          
                                    //без инициализации xls файлом
};

#endif

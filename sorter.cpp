#include "sorter.h"
#include <string.h>

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
Error::text()
{
    return mess;
}

Sorter::Sorter(char* filename)
{
    //открываем файл лалала
}

#include <stdio.h>
#include <locale.h>
#include <stdlib.h>

void main(int argc, char *argv[])
{
    setlocale(LC_ALL, "Rus");
    printf("Говорим по-русски\n");
    
    if (argc > 1) {
        int val = atoi(argv[1]);
        printf("Нулевой параметр: %s\n", argv[0]);
        printf("Первый параметр: %s\n", argv[1]);
        printf("Преобразованное число: %d\n", val);
    } else {
        printf("Параметры не переданы.\n");
    }
}

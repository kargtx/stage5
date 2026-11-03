// ЛР3, раздел 3. Программа «Всем привет!» (рис. 3.1)
// Сборка: gcc -O2 mpi1-hello.c -I"%MSMPI_INC%" -L"%MSMPI_LIB64%" -lmsmpi -o mpi1-hello.exe
// Запуск: mpi1-hello            (последовательно)
//         mpiexec -n 4 mpi1-hello (параллельно)
#include <stdio.h>
#include <mpi.h>

int main()
{
    MPI_Init(NULL, NULL);
    printf("Hello, World!\n");
    MPI_Finalize();
    return 0;
}

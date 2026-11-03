// ЛР3, раздел 5. Параметр командной строки (рис. 5.1)
// Запуск: mpiexec -n 2 mpi3-arg 100
#include <stdio.h>
#include <stdlib.h>
#include <mpi.h>

int main(int argc, char* argv[])
{
    int ProcessRank;
    long long N;
    MPI_Init(&argc, &argv);
    MPI_Comm_rank(MPI_COMM_WORLD, &ProcessRank);
    N = atoll(argv[1]);                 // строка -> длинное целое (8 байт)
    printf("Process %d:\t", ProcessRank);
    printf("N = %lld\n", N);
    MPI_Finalize();
    return 0;
}

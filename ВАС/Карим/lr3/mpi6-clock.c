// ЛР3, раздел 8. Замер времени параллельного суммирования (рис. 8.1)
// Вывод: число процессов <TAB> время, с <TAB> сумма
// Запуск: mpiexec -n 4 mpi6-clock 1000000000
#include <stdio.h>
#include <stdlib.h>
#include <time.h>
#include <mpi.h>

int main(int argc, char* argv[])
{
    int ProcessRank, CommunicatorSize;
    long long N, i, Z = 0;
    volatile long long S = 0;
    double T;
    clock_t t0, t1;
    MPI_Init(&argc, &argv);
    N = atoll(argv[1]);
    MPI_Comm_rank(MPI_COMM_WORLD, &ProcessRank);
    MPI_Comm_size(MPI_COMM_WORLD, &CommunicatorSize);
    t0 = clock();
    for (i = ProcessRank; i < N; i += CommunicatorSize)
        S++;
    long long Sl = S;
    MPI_Reduce(&Sl, &Z, 1, MPI_LONG_LONG, MPI_SUM, 0, MPI_COMM_WORLD);
    if (ProcessRank == 0)
    {
        t1 = clock();
        T = (double)(t1 - t0) / CLOCKS_PER_SEC;
        printf("%d\t%f\t%lld\n", CommunicatorSize, T, Z);
    }
    MPI_Finalize();
    return 0;
}

// ЛР3, раздел 6. Независимое суммирование единиц в каждом процессе (рис. 6.1)
// Запуск: mpiexec -n 8 mpi4-sum 1000000000
// volatile - чтобы оптимизатор gcc (-O2) не заменил цикл одной операцией S = N
#include <stdio.h>
#include <stdlib.h>
#include <mpi.h>

int main(int argc, char* argv[])
{
    int ProcessRank;
    long long N, i;
    volatile long long S = 0;
    MPI_Init(&argc, &argv);
    N = atoll(argv[1]);
    MPI_Comm_rank(MPI_COMM_WORLD, &ProcessRank);
    for (i = 0; i < N; i++)
        S++;
    printf("%d:\t", ProcessRank);
    printf("%lld\t%lld\n", N, S);
    MPI_Finalize();
    return 0;
}

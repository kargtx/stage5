// ЛР3, раздел 7. Раздача заданий чередованием и редукция MPI_Reduce (рис. 7.1)
// Запуск: mpiexec -n 8 mpi5-reduce 1000
#include <stdio.h>
#include <stdlib.h>
#include <mpi.h>

int main(int argc, char* argv[])
{
    int ProcessRank, CommunicatorSize;
    long long N, i, Z = 0;
    volatile long long S = 0;
    MPI_Init(&argc, &argv);
    N = atoll(argv[1]);
    MPI_Comm_rank(MPI_COMM_WORLD, &ProcessRank);
    MPI_Comm_size(MPI_COMM_WORLD, &CommunicatorSize);
    for (i = ProcessRank; i < N; i += CommunicatorSize)   // метод чередования
        S++;
    printf("S[%d] = %lld\n", ProcessRank, S);
    long long Sl = S;
    MPI_Reduce(&Sl, &Z, 1, MPI_LONG_LONG, MPI_SUM, 0, MPI_COMM_WORLD);
    if (ProcessRank == 0)
        printf("%lld\t%lld\n", N, Z);
    MPI_Finalize();
    return 0;
}

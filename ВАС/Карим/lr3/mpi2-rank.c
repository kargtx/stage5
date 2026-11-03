// ЛР3, раздел 4. Номер процесса и размер коммуникатора (рис. 4.1)
// Запуск: mpiexec -n 4 mpi2-rank
#include <stdio.h>
#include <mpi.h>

int main()
{
    int ProcessRank, CommunicatorSize;
    MPI_Init(NULL, NULL);
    MPI_Comm_rank(MPI_COMM_WORLD, &ProcessRank);       // номер процесса в группе
    MPI_Comm_size(MPI_COMM_WORLD, &CommunicatorSize);  // число процессов в группе
    printf("Process %d of %d\n", ProcessRank, CommunicatorSize);
    MPI_Finalize();
    return 0;
}

// ЛР3, раздел 10. Параллельное численное интегрирование, одинарная точность (рис. 10.1)
// Вариант 4: f(x) = 0.05*x^3 + 0.3*x^2 - 20*x + 200,  x in [-5; 20]
// Точное значение: 4054.6875
// Метод средних прямоугольников, раздача заданий чередованием, сбор MPI_Reduce
// Вывод: p <TAB> N <TAB> T, с <TAB> интеграл
// Запуск: mpiexec -n 4 mpi7-integral 100000000
#include <stdio.h>
#include <stdlib.h>
#include <time.h>
#include <mpi.h>

#define A (-5.0)
#define B 20.0

int main(int argc, char* argv[])
{
    int ProcessRank, CommunicatorSize;
    long long N, i;
    float T, dx, x, Z = 0.0, S = 0.0;
    MPI_Init(&argc, &argv);
    N = atoll(argv[1]);
    MPI_Comm_rank(MPI_COMM_WORLD, &ProcessRank);
    MPI_Comm_size(MPI_COMM_WORLD, &CommunicatorSize);
    clock_t t0 = clock();
    dx = (B - A) / N;                                   // ширина прямоугольника
    for (i = ProcessRank; i < N; i += CommunicatorSize)
    {
        x = A + (i + 0.5) * dx;                         // середина прямоугольника
        S = S + 0.05 * x * x * x + 0.3 * x * x - 20.0 * x + 200.0;
    }
    MPI_Reduce(&S, &Z, 1, MPI_FLOAT, MPI_SUM, 0, MPI_COMM_WORLD);
    MPI_Finalize();
    if (ProcessRank == 0)
    {
        Z = Z * dx;
        clock_t t1 = clock();
        T = (float)(t1 - t0) / CLOCKS_PER_SEC;
        printf("%d\t%lld\t%.10f\t%.10f\n", CommunicatorSize, N, T, Z);
    }
    return 0;
}

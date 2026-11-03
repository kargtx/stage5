// ЛР3, раздел 11.1 (повышенной сложности). Гипертрединг: привязка процессов к процессорам
// Интегрирование (double) с привязкой каждого MPI-процесса к логическому процессору
// через SetThreadAffinityMask.
//   Mode = 0 - без привязки (решает ОС)
//   Mode = 1 - привязка ко всем логическим процессорам подряд: процесс k -> ЛП k
//              (4 процесса = 2 физических ядра по 2 потока HT)
//   Mode = 2 - привязка только к физическим ядрам:   процесс k -> ЛП 2k
//              (4 процесса = 4 физических ядра, второй поток HT не используется)
// Вывод: p <TAB> Mode <TAB> N <TAB> T, с <TAB> интеграл
// Запуск: mpiexec -n 4 mpi9-affinity 1000000000 2
#include <windows.h>
#include <stdio.h>
#include <stdlib.h>
#include <time.h>
#include <mpi.h>

#define A (-5.0)
#define B 20.0

int main(int argc, char* argv[])
{
    int ProcessRank, CommunicatorSize, Mode;
    long long N, i;
    double T, dx, x, Z = 0.0, S = 0.0;
    MPI_Init(&argc, &argv);
    N = atoll(argv[1]);
    Mode = (argc > 2) ? atoi(argv[2]) : 0;
    MPI_Comm_rank(MPI_COMM_WORLD, &ProcessRank);
    MPI_Comm_size(MPI_COMM_WORLD, &CommunicatorSize);

    if (Mode == 1)
        SetThreadAffinityMask(GetCurrentThread(), (DWORD_PTR)1 << ProcessRank);
    else if (Mode == 2)
        SetThreadAffinityMask(GetCurrentThread(), (DWORD_PTR)1 << (2 * ProcessRank));

    MPI_Barrier(MPI_COMM_WORLD);       // общий старт всех процессов
    clock_t t0 = clock();
    dx = (B - A) / N;
    for (i = ProcessRank; i < N; i += CommunicatorSize)
    {
        x = A + (i + 0.5) * dx;
        S = S + 0.05 * x * x * x + 0.3 * x * x - 20.0 * x + 200.0;
    }
    MPI_Reduce(&S, &Z, 1, MPI_DOUBLE, MPI_SUM, 0, MPI_COMM_WORLD);
    if (ProcessRank == 0)
    {
        Z = Z * dx;
        clock_t t1 = clock();
        T = (double)(t1 - t0) / CLOCKS_PER_SEC;
        printf("%d\t%d\t%lld\t%.4f\t%.10f\n", CommunicatorSize, Mode, N, T, Z);
    }
    MPI_Finalize();
    return 0;
}

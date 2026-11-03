@echo off
rem 11.1 Гипертрединг: 4 процесса, режимы 0 (ОС), 1 (ЛП 0-3), 2 (ЛП 0,2,4,6), по 10 прогонов
echo p	Mode	N	T	S
for /L %%k in (1, 1, 10) do for %%m in (0 1 2) do mpiexec -n 4 mpi9-affinity 2000000000 %%m

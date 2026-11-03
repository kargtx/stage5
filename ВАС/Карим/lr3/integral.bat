@echo off
rem %1 = имя программы (mpi7-integral или mpi8-integral-double)
rem Серия 1: точность - N = 10^3 ... 10^9, p = 1, 2, 4, 8, по 10 прогонов
echo p	N	T	S
for /L %%k in (1, 1, 10) do for %%n in (1000 10000 100000 1000000 10000000 100000000 1000000000) do for %%i in (1 2 4 8) do mpiexec -n %%i %1 %%n

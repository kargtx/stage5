@echo off
rem %1 = имя программы; Серия 2: ускорение - N = 10^9, p = 1..8, по 10 прогонов
echo p	N	T	S
for /L %%k in (1, 1, 10) do for /L %%i in (1, 1, 8) do mpiexec -n %%i %1 1000000000

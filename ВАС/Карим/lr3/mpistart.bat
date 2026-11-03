@echo off
echo p	t	S
for /L %%k in (1, 1, 10) do for /L %%i in (1, 1, 8) do mpiexec -n %%i mpi6-clock 1000000000

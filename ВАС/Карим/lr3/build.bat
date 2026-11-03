@echo off
rem Сборка всех программ ЛР3 компилятором gcc (MinGW-w64) с библиотекой Microsoft MPI

cd /d "%~dp0"
set "MPIINC=C:\Program Files (x86)\Microsoft SDKs\MPI\Include"
set "MPILIB=C:\Program Files (x86)\Microsoft SDKs\MPI\Lib\x64"
where gcc
echo.
for %%f in (mpi1-hello mpi2-rank mpi3-arg mpi4-sum mpi5-reduce mpi6-clock mpi7-integral mpi8-integral-double mpi9-affinity) do (
  echo gcc -O2 %%f.c -I"%%MPIINC%%" -L"%%MPILIB%%" -lmsmpi -o %%f.exe
  gcc -O2 -Wall %%f.c -I"%MPIINC%" -L"%MPILIB%" -lmsmpi -o %%f.exe
)
echo.
dir /b *.exe
pause

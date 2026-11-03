@echo off
cd /d "%~dp0"
echo ==== gcc > check.txt
where gcc >> check.txt 2>&1
gcc --version >> check.txt 2>&1
echo ==== mpiexec >> check.txt
where mpiexec >> check.txt 2>&1
where smpd >> check.txt 2>&1
echo ==== MSMPI env >> check.txt
set MSMPI >> check.txt 2>&1
echo ==== dirs >> check.txt
dir /b "C:\Program Files\Microsoft MPI\Bin" >> check.txt 2>&1
dir /s /b "C:\Program Files (x86)\Microsoft SDKs\MPI" >> check.txt 2>&1
dir /b "C:\Program Files (x86)\MPICH2" >> check.txt 2>&1
dir /b "C:\Program Files\MPICH2" >> check.txt 2>&1
dir /b C:\msys64 >> check.txt 2>&1
dir /b C:\MinGW >> check.txt 2>&1
dir /b C:\mingw64 >> check.txt 2>&1
dir /b C:\msys64\mingw64\include\mpi.h >> check.txt 2>&1
dir /b C:\msys64\ucrt64\include\mpi.h >> check.txt 2>&1
echo ==== PATH >> check.txt
echo %PATH% >> check.txt
echo ==== cpu >> check.txt
powershell -NoProfile -Command "Get-CimInstance Win32_Processor | Select Name,NumberOfCores,NumberOfLogicalProcessors | Format-List" >> check.txt 2>&1
type check.txt
pause

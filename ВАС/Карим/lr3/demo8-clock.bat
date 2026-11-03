@chcp 65001 >nul
@cd /d "%~dp0"
@set "PATH=%PATH%;C:\Program Files\Microsoft MPI\Bin"
@echo on
mpi6-clock 1000000000
mpiexec -n 2 mpi6-clock 1000000000
mpiexec -n 4 mpi6-clock 1000000000
mpiexec -n 8 mpi6-clock 1000000000
@pause

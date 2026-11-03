@chcp 65001 >nul
@cd /d "%~dp0"
@set "PATH=%PATH%;C:\Program Files\Microsoft MPI\Bin"
@echo on
mpi4-sum 1000000
mpiexec -n 2 mpi4-sum 1000000
mpiexec -n 4 mpi4-sum 1000000
mpiexec -n 8 mpi4-sum 100000000
@pause

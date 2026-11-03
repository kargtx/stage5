@chcp 65001 >nul
@cd /d "%~dp0"
@set "PATH=%PATH%;C:\Program Files\Microsoft MPI\Bin"
@echo on
mpi2-rank
mpiexec -n 4 mpi2-rank
mpiexec -n 4 mpi2-rank
mpiexec -n 8 mpi2-rank
@pause

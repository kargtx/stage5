@chcp 65001 >nul
@cd /d "%~dp0"
@set "PATH=%PATH%;C:\Program Files\Microsoft MPI\Bin"
@echo on
mpiexec -n 8 mpi5-reduce 1000
mpiexec -n 8 mpi5-reduce 1000
mpiexec -n 3 mpi5-reduce 1000
@pause

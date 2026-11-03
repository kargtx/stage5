@chcp 65001 >nul
@cd /d "%~dp0"
@set "PATH=%PATH%;C:\Program Files\Microsoft MPI\Bin"
@echo on
mpi1-hello
mpiexec -n 2 mpi1-hello
mpiexec -n 4 mpi1-hello
mpiexec -n 6 mpi1-hello
@pause

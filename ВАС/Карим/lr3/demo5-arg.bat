@chcp 65001 >nul
@cd /d "%~dp0"
@set "PATH=%PATH%;C:\Program Files\Microsoft MPI\Bin"
@echo on
mpi3-arg 100
mpi3-arg 9223372036854775807
mpi3-arg 9223372036854775808
mpi3-arg -9223372036854775809
mpiexec -n 2 mpi3-arg 100
mpiexec -n 4 mpi3-arg 12345
@pause

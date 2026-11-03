@chcp 65001 >nul
@cd /d "%~dp0"
@set "PATH=%PATH%;C:\Program Files\Microsoft MPI\Bin"
@echo on
mpi7-integral 100000000
mpiexec -n 2 mpi7-integral 100000000
mpiexec -n 4 mpi7-integral 100000000
mpiexec -n 8 mpi7-integral 100000000
mpi8-integral-double 100000000
mpiexec -n 8 mpi8-integral-double 100000000
@pause

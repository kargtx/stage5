@chcp 65001 >nul
@cd /d "%~dp0"
@set "PATH=%PATH%;C:\Program Files\Microsoft MPI\Bin"
@echo on
mpiexec -n 8 mpi4-sum 30000000000
@pause

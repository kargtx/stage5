@cd /d "%~dp0"
@set "PATH=%PATH%;C:\Program Files\Microsoft MPI\Bin"
@echo off
echo [%time%] 1/6 mpistart - section 9          & call mpistart.bat > m.txt
echo [%time%] 2/6 float  - accuracy             & call integral.bat mpi7-integral > float_acc.txt
echo [%time%] 3/6 float  - speedup              & call speedup.bat  mpi7-integral > float_speed.txt
echo [%time%] 4/6 double - accuracy             & call integral.bat mpi8-integral-double > double_acc.txt
echo [%time%] 5/6 double - speedup              & call speedup.bat  mpi8-integral-double > double_speed.txt
echo [%time%] 6/6 hyperthreading                & call hyper.bat > hyper.txt
echo [%time%] DONE > done.txt
echo [%time%] DONE

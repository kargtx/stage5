@echo off
echo Compiling C programs...
gcc 3_1_hello.c -o 3_1_hello.exe
gcc 3_3_args.c -o 3_3_args.exe
gcc 4_1_sum_seq.c -o 4_1_sum_seq.exe
gcc 5_1_thread.c -o 5_1_thread.exe
gcc 5_3_sum_par_race.c -o 5_3_sum_par_race.exe
gcc 5_4_sum_par_cs.c -o 5_4_sum_par_cs.exe
gcc 5_5_sum_par_fast.c -o 5_5_sum_par_fast.exe
gcc 6_5_integ_seq.c -o 6_5_integ_seq.exe
gcc 6_7_integ_par.c -o 6_7_integ_par.exe
echo Done.

@echo off
echo Running experiments...

echo --- 4.1 Sequential Sum ---
echo S,Ns,dT > results_4_1.csv
4_1_sum_seq.exe 1000000000 >> results_4_1.csv
4_1_sum_seq.exe 1000000000 >> results_4_1.csv
4_1_sum_seq.exe 1000000000 >> results_4_1.csv

echo --- 5.3 Parallel Sum (Race) ---
echo S,Ns,NTh,Error,dT > results_5_3.csv
5_3_sum_par_race.exe 100000000 >> results_5_3.csv
5_3_sum_par_race.exe 100000000 >> results_5_3.csv
5_3_sum_par_race.exe 100000000 >> results_5_3.csv

echo --- 5.4 Parallel Sum (Critical Section) ---
echo S,Ns,NTh,Error,dT > results_5_4.csv
5_4_sum_par_cs.exe 100000 >> results_5_4.csv
5_4_sum_par_cs.exe 100000 >> results_5_4.csv
5_4_sum_par_cs.exe 100000 >> results_5_4.csv

echo --- 5.5 Parallel Sum (Fast) ---
echo S,Ns,NTh,Error,dT > results_5_5.csv
5_5_sum_par_fast.exe 100000000 >> results_5_5.csv
5_5_sum_par_fast.exe 100000000 >> results_5_5.csv
5_5_sum_par_fast.exe 100000000 >> results_5_5.csv

echo --- 6.5 Sequential Integration ---
echo N,S,Error_pct,dT > results_6_5.csv
6_5_integ_seq.exe 10000 >> results_6_5.csv
6_5_integ_seq.exe 100000 >> results_6_5.csv
6_5_integ_seq.exe 1000000 >> results_6_5.csv
6_5_integ_seq.exe 10000000 >> results_6_5.csv
6_5_integ_seq.exe 100000000 >> results_6_5.csv

echo --- 6.7 Parallel Integration ---
echo N,NTh,S,Error_pct,dT > results_6_7.csv
6_7_integ_par.exe 100000000 1 >> results_6_7.csv
6_7_integ_par.exe 100000000 2 >> results_6_7.csv
6_7_integ_par.exe 100000000 4 >> results_6_7.csv
6_7_integ_par.exe 100000000 8 >> results_6_7.csv

echo All experiments finished.

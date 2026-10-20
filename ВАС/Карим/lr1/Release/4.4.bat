@echo off
echo S,Ns,dT > results.csv
lr1.exe 1000000000 >> results.csv
lr1.exe 5000000000 >> results.csv
lr1.exe 10000000000 >> results.csv
lr1.exe 22000000000 >> results.csv

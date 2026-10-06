#include <windows.h>
#include <stdio.h>
#include <time.h>
#include <stdlib.h>

#define NTh 20

int S = 0;
int Ns = 10000000;

void WINAPI ThreadProc()
{
    for (int i = 0; i < Ns; i++)
        S = S + 1;
}

void main(int argc, char *argv[])
{
    if (argc > 1) {
        Ns = atoi(argv[1]);
    }
    
    HANDLE Thread[NTh];
    clock_t t1, t2;
    float dT;
    
    t1 = clock();
    for (int i = 0; i < NTh; i++)
        Thread[i] = CreateThread(0, 0, (LPTHREAD_START_ROUTINE)ThreadProc, 0, 0, 0);
        
    WaitForMultipleObjects(NTh, Thread, TRUE, INFINITE);
    t2 = clock();
    
    dT = (float)(t2 - t1) / CLOCKS_PER_SEC;
    
    int expected_S = Ns * NTh;
    int error = expected_S - S;
    
    printf("%d,%d,%d,%d,%5.10f\n", S, Ns, NTh, error, dT);
}

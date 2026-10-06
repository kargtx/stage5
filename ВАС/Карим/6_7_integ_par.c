#include <windows.h>
#include <stdio.h>
#include <time.h>
#include <stdlib.h>

#define MAX_THREADS 64

float A = -5.0f;
float B = 20.0f;
int N = 100000;
int num_threads = 4;
float global_S = 0.0f;
CRITICAL_SECTION CS;

float f(float x)
{
    return 0.05f * x * x * x + 0.3f * x * x - 20.0f * x + 200.0f;
}

struct ThreadData {
    int id;
};

void WINAPI ThreadProc(LPVOID lpParam)
{
    struct ThreadData* data = (struct ThreadData*)lpParam;
    int id = data->id;
    
    // Set affinity to the specific core
    SetThreadAffinityMask(GetCurrentThread(), (DWORD_PTR)1 << id);
    
    float dx = (B - A) / N;
    float local_S = 0.0f;
    
    // Interleaved distribution
    for (int i = id; i < N; i += num_threads) {
        float x = A + dx * (i + 0.5f);
        local_S += f(x) * dx;
    }
    
    EnterCriticalSection(&CS);
    global_S += local_S;
    LeaveCriticalSection(&CS);
}

void main(int argc, char *argv[])
{
    if (argc > 1) N = atoi(argv[1]);
    if (argc > 2) num_threads = atoi(argv[2]);
    if (num_threads > MAX_THREADS) num_threads = MAX_THREADS;
    
    HANDLE Thread[MAX_THREADS];
    struct ThreadData tData[MAX_THREADS];
    clock_t t1, t2;
    float dT;
    
    InitializeCriticalSection(&CS);
    
    t1 = clock();
    for (int i = 0; i < num_threads; i++) {
        tData[i].id = i;
        Thread[i] = CreateThread(0, 0, (LPTHREAD_START_ROUTINE)ThreadProc, &tData[i], 0, 0);
    }
        
    WaitForMultipleObjects(num_threads, Thread, TRUE, INFINITE);
    t2 = clock();
    
    DeleteCriticalSection(&CS);
    
    dT = (float)(t2 - t1) / CLOCKS_PER_SEC;
    
    float exact = 4054.6875f;
    float err = exact - global_S;
    float err_pct = (err / exact) * 100.0f;
    if (err_pct < 0) err_pct = -err_pct;
    
    printf("%d,%d,%f,%f,%5.10f\n", N, num_threads, global_S, err_pct, dT);
}

#include <stdio.h>
#include <time.h>
#include <stdlib.h>

float f(float x)
{
    return 0.05f * x * x * x + 0.3f * x * x - 20.0f * x + 200.0f;
}

void main(int argc, char *argv[])
{
    int N = 100; // default intervals
    if (argc > 1) {
        N = atoi(argv[1]);
    }
    
    float A = -5.0f;
    float B = 20.0f;
    float dx = (B - A) / N;
    float S = 0.0f;
    
    clock_t t1, t2;
    float dT;
    
    t1 = clock();
    for (int i = 0; i < N; i++) {
        float x = A + dx * (i + 0.5f);
        S += f(x) * dx;
    }
    t2 = clock();
    
    dT = (float)(t2 - t1) / CLOCKS_PER_SEC;
    
    float exact = 4054.6875f;
    float err = exact - S;
    float err_pct = (err / exact) * 100.0f;
    if (err_pct < 0) err_pct = -err_pct;
    
    printf("%d,%f,%f,%5.10f\n", N, S, err_pct, dT);
}

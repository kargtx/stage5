#include <stdio.h>
#include <time.h>
#include <stdlib.h>

void main(int argc, char *argv[])
{
    int Ns = 1000000;
    if (argc > 1) {
        Ns = atoi(argv[1]);
    }
    
    int S = 0;
    int i;
    float dT;
    clock_t t1, t2;
    
    t1 = clock();
    for (i = 0; i < Ns; i++) {
        S = S + 1;
    }
    t2 = clock();
    
    dT = (float)(t2 - t1) / CLOCKS_PER_SEC;
    
    // Output separated by commas to easily import in CSV
    printf("%d,%d,%5.10f\n", S, Ns, dT);
}

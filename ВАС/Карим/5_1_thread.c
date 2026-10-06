#include <windows.h>
#include <stdio.h>

void WINAPI Thread2()
{
    printf("Thread Two\n");
    // while(1); // commented out to allow exit
}

void main()
{
    CreateThread(0, 0, (LPTHREAD_START_ROUTINE)Thread2, 0, 0, 0);
    printf("Main - Thread One\n");
    getchar(); // Wait for input to see threads
}

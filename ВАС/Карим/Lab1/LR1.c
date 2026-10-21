
// Задание 3.1. Программа "Hello, World!" 
/*
#include <stdio.h>

void main() { 
    
    printf("Hello, World\n");
    //getchar(); // Ожидание нажатия Enter
}

*/
// 
// Задания 3.3, 3.4, 3.5. Параметры командной строки (argc, argv), 
// локализация стандартного вывода (setlocale) и преобразование строки в число (atoi)
// 
//3.3
/*
#include <stdio.h>

void main(int argc, char* argv[])
{


    // Выводим все полученные параметры
    for (int i = 0; i < argc; i++) {
        printf("Параметр [%d]: %s\n", i, argv[i]);
    }
    //d - целое число, s - строка
}
*/

//3.4
/*
#include <stdio.h>
#include <locale.h>

void main(int argc, char* argv[])
{
    setlocale(LC_ALL, "Russian"); 
    for (int i = 0; i < argc; i++) {
        printf("Параметр [%d]: %s\n", i, argv[i]);
    }
}
*/

//3.5
/*
#include <stdlib.h>.
#include <stdio.h>
#include <locale.h>

void main(int argc, char* argv[])
{
    setlocale(LC_ALL, "Russian");
    //преобразовываем элемент с индексом 1 из строки в число
    int n = atoi(argv[1]);
    for (int i = 0; i < argc; i++) {
        printf("Параметр [%d]: %s\n", i, argv[i]);
    }

    printf("Преобразованное число: %d\n", n);
    printf("Умноженное на 2 число: %d\n", n*2);

}
*/

// Задания 4.1, 4.2. Последовательное суммирование единиц и замер времени работы
// Измерение времени с помощью функции clock() и вывод в формате CSV для Excel

//4.1
/*
#include <stdlib.h>.
#include <stdio.h>
#include <locale.h>

void main(int argc, char* argv[])
{
    setlocale(LC_ALL, "Russian");

    //int Ns = atoi(argv[1]);
    int Ns = 2147483648;
    int S = 0;
    int i;
    

    for (i = 0; i < Ns; i++)
        S = S + 1;
    printf("Заданное значение = %d\n", Ns);
    printf("S = %d\n", S);
    printf("Ns = %d\n", Ns);
}
*/

//4.2
/*
#include <stdlib.h>.
#include <stdio.h>
#include <time.h>    
#include <locale.h>

void main(int argc, char* argv[])
{
    setlocale(LC_ALL, "Russian");

    volatile long long S = 0;
    long long Ns = _atoi64(argv[1]);
    long long i;
    float dT;
    clock_t t1, t2;

    // Засекаем начальное время
    t1 = clock();

    for (i = 0; i < Ns; i++)
        S = S + 1;

    // Засекаем конечное время
    t2 = clock();

    // Переводим разницу тактов в секунды
    dT = (float)(t2 - t1) / CLOCKS_PER_SEC;

    printf("S = %lld\n", S);
    printf("Ns = %lld\n", Ns);
    printf("Время работы dT = %5.10f сек\n", dT);
}
*/

//4.4
/*
#include <stdlib.h>
#include <stdio.h>
#include <time.h>
#include <windows.h>

void main(int argc, char* argv[])
{
    SetProcessAffinityMask(GetCurrentProcess(), 255); // 1 в двоичной: 0001
    volatile long long S = 0;
    long long Ns = _atoi64(argv[1]);
    

    long long i;
    float dT;
    clock_t t1, t2;

    t1 = clock();
    for (i = 0; i < Ns; i++) {
        S = S + 1;
    }
    t2 = clock();

    dT = (float)(t2 - t1) / CLOCKS_PER_SEC;

    // Вывод в одну строку через запятую: S, Ns, dT
    printf("%lld,%lld,%5.10f\n", S, Ns, dT);
}
*/

// Задание 5.1. Создание параллельного потока с помощью функции Windows API CreateThread()
 
/*
#include <windows.h>
#include <stdio.h>

void WINAPI Thread2()
{
    printf("Thread Two\n");
    while(1); // закомментировано для корректного выхода
}

void main()
{
    CreateThread(0, 0, (LPTHREAD_START_ROUTINE)Thread2, 0, 0, 0);
    printf("Main - Thread One\n");
    getchar(); // Ожидание нажатия клавиши для просмотра созданных потоков в Диспетчере задач
}
*/



// Задание 5.3. Параллельные потоки и ситуация гонки данных (Data Race)

/*
#include <windows.h>
#include <stdio.h>
#include <time.h>
#define NTh 20
#define Ns 10000000

int S = 0;

void WINAPI ThreadProc()
{
    for (int i = 0; i < Ns; i++)
        S = S + 1;
}

void main()
{
    HANDLE Thread[NTh];
    clock_t t1, t2;
    t1 = clock();
    for (int i = 0; i < NTh; i++)
        Thread[i] = CreateThread(0, 0, (LPTHREAD_START_ROUTINE)ThreadProc, 0, 0, 0);
    WaitForMultipleObjects(NTh, Thread, TRUE, INFINITE);
    t2 = clock();
    printf("S = %d\nError = %d\n", S, Ns * NTh - S);
    printf("T = %5.10f\n", (double)(t2 - t1) / CLOCKS_PER_SEC);

}
*/



// Задание 5.4. Синхронизация потоков с помощью критической секции (Critical Section)
// Блокировка каждого обращения к общей переменной S (надежно, но медленно)

/*
#include <windows.h>
#include <stdio.h>
#include <time.h>
#define NTh 20
#define Ns 1000000

int S = 0;
CRITICAL_SECTION CS;

void WINAPI ThreadProc()
{
    int i;
    for (i = 0; i < Ns; i++)
    {
        EnterCriticalSection(&CS);
        S = S + 1;
        LeaveCriticalSection(&CS);
    }
}

void main()
{
    int i;
    HANDLE Thread[NTh];
    clock_t t1, t2;
    InitializeCriticalSection(&CS);
    t1 = clock();
    for (i = 0; i < NTh; i++)
        Thread[i] = CreateThread(0, 0, (LPTHREAD_START_ROUTINE)ThreadProc, 0, 0, 0);
    WaitForMultipleObjects(NTh, Thread, TRUE, INFINITE);
    t2 = clock();
    DeleteCriticalSection(&CS);
    printf("S = %d\nError = %d\n", S, Ns * NTh - S);
    printf("T = %5.10f\n", (double)(t2 - t1) / CLOCKS_PER_SEC);
}
*/


// Задание 5.5. Оптимизация взаимодействия потоков через локальные переменные
// Потоки накапливают сумму локально, а в критическую секцию входят лишь один раз

/*
#include <windows.h>
#include <stdio.h>
#include <time.h>
#define NTh 20
#define Ns 100000000LL

long long S = 0;
CRITICAL_SECTION CS;

void WINAPI ThreadProc()
{
    long long i, local = 0;
    for (i = 0; i < Ns; i++)
        local = local + 1;
    EnterCriticalSection(&CS);
    S = S + local;
    LeaveCriticalSection(&CS);
}

void main()
{
    int i;
    HANDLE Thread[NTh];
    clock_t t1, t2;
    InitializeCriticalSection(&CS);
    t1 = clock();
    for (i = 0; i < NTh; i++)
        Thread[i] = CreateThread(0, 0, (LPTHREAD_START_ROUTINE)ThreadProc, 0, 0, 0);
    WaitForMultipleObjects(NTh, Thread, TRUE, INFINITE);
    t2 = clock();
    DeleteCriticalSection(&CS);
    printf("S = %lld\nError = %lld\n", S, Ns * NTh - S);
    printf("T = %5.10f\n", (double)(t2 - t1) / CLOCKS_PER_SEC);
}
*/

// Раздел 5.6. Оценка эффективности распараллеливания

/*
#include <windows.h>
#include <stdio.h>
#include <time.h>
#define NTh 20
#define Ns 100000000LL

long long S = 0;
CRITICAL_SECTION CS;

void WINAPI ThreadProc()
{
    long long i, local = 0;
    for (i = 0; i < Ns; i++)
        local = local + 1;
    EnterCriticalSection(&CS);
    S = S + local;
    LeaveCriticalSection(&CS);
}

int main()
{
    int i;
    HANDLE Thread[NTh];
    clock_t t1, t2;
    InitializeCriticalSection(&CS);
    t1 = clock();
    for (i = 0; i < NTh; i++)
        Thread[i] = CreateThread(0, 0, (LPTHREAD_START_ROUTINE)ThreadProc, 0, 0, 0);
    WaitForMultipleObjects(NTh, Thread, TRUE, INFINITE);
    t2 = clock();
    DeleteCriticalSection(&CS);
    printf("S = %lld\nError = %lld\n", S, Ns * NTh - S);
    printf("T = %5.10f\n", (double)(t2 - t1) / CLOCKS_PER_SEC);
    return 0;
}
*/


// 
// Раздел 6.5. Численное интегрирование методом прямоугольников (вариант 4)
// Функция: f(x) = 0.05*x^3 + 0.3*x^2 - 20*x + 200 на [-5, 20]
// Аналитическое значение: S_exact = 4054.6875
// 
/*
#include <stdio.h>
#include <stdlib.h>
#include <time.h>

#define A (-5.0)
#define B 20.0
#define S_EXACT 4054.6875

int main(int argc, char *argv[])
{
    long long i, N;
    float dx, x, f, sum = 0.0f, S;
    clock_t t1, t2;
    double T;

    N = (argc > 1) ? atoll(argv[1]) : 100;
    dx = (B - A) / N;

    t1 = clock();
    for (i = 0; i < N; i++) {
        x = A + (i + 0.5f) * dx; // середина очередного отрезка
        f = 0.05f * x * x * x + 0.3f * x * x - 20.0f * x + 200.0f; // f(x) вариант 4 // вычисление высоты
        sum = sum + f; // все высоты отрезков
    }
    S = sum * dx;
    t2 = clock();
    T = (double)(t2 - t1) / CLOCKS_PER_SEC;

    printf("N = %lld\n", N);
    printf("S = %.6f\n", S);
    printf("S_exact = %.6f\n", S_EXACT);
    printf("Error = %.6f\n", S - S_EXACT);
    printf("T = %.4f\n", T);
    return 0;
}
*/


// 
// Раздел 6.6. Тип данных float / double
// Часть 1: пределы типов
// 
/*
#include <stdio.h>
#include <float.h>

int main()
{
    float  f;
    double d;

    printf("float:  %d bytes, %d digits, min %e, max %e\n",
           (int)sizeof(float), FLT_DIG, FLT_MIN, FLT_MAX);
    printf("double: %d bytes, %d digits, min %e, max %e\n\n",
           (int)sizeof(double), DBL_DIG, DBL_MIN, DBL_MAX);

    f = FLT_MAX;  f = f * 10;    printf("float  max * 10   = %e\n", f);
    f = -FLT_MAX; f = f * 10;    printf("float -max * 10   = %e\n", f);
    f = FLT_MIN;  f = f / 1e10f; printf("float  min / 1e10 = %e\n", f);

    d = DBL_MAX;  d = d * 10;    printf("double  max * 10   = %e\n", d);
    d = -DBL_MAX; d = d * 10;    printf("double -max * 10   = %e\n", d);
    d = DBL_MIN;  d = d / 1e20;  printf("double  min / 1e20 = %e\n\n", d);

    f = 100000000.0f; f = f + 1.0f;
    printf("float  1e8 + 1  = %.1f\n", f);
    d = 100000000.0;  d = d + 1.0;
    printf("double 1e8 + 1  = %.1f\n", d);
    return 0;
}
*/

// 
// Раздел 6.6. Численное интегрирование: float vs double (rect2.c, вариант 4)
// Функция: f(x) = 0.05*x^3 + 0.3*x^2 - 20*x + 200 на [-5, 20]
// Компилировать как:
//   gcc rect2.c -o rect_f.exe
//   gcc -DREAL=double rect2.c -o rect_d.exe
// 
/*
#include <stdio.h>
#include <stdlib.h>
#include <time.h>

#ifndef REAL
#define REAL float
#endif

#define A (-5.0)
#define B 20.0
#define S_EXACT 4054.6875

int main(int argc, char *argv[])
{
    long long i, N;
    REAL dx, x, f, sum = 0, S;
    clock_t t1, t2;
    double T, Err;

    if (argc < 2)
        return 1;
    N = atoll(argv[1]);
    dx = (B - A) / N;

    t1 = clock();
    for (i = 0; i < N; i++) {
        x = A + (i + 0.5) * dx;
        f = 0.05 * x * x * x + 0.3 * x * x - 20.0 * x + 200.0; // f(x) вариант 4
        sum = sum + f;
    }
    S = sum * dx;
    t2 = clock();

    T = (double)(t2 - t1) / CLOCKS_PER_SEC;
    Err = (S - S_EXACT) / S_EXACT * 100;
    printf("%lld,%.10f,%.4f,%e\n", N, (double)S, T, Err);
    return 0;
}

// gcc rect2.c -o rect_f.exe
// gcc -DREAL=double rect2.c -o rect_d.exe
*/


// 
// Раздел 6.7. Параллельное численное интегрирование с SetThreadAffinityMask (integ.c)
// Функция: f(x) = 0.05*x^3 + 0.3*x^2 - 20*x + 200 на [-5, 20]
// Аналитическое значение: S_exact = 4054.6875
// Запуск: integ.exe N NTh [pin] [wait]
// 

#include <windows.h>
#include <stdio.h>
#include <stdlib.h>
#include <time.h>

#ifndef REAL
#define REAL double
#endif

#define A (-5.0)
#define B 20.0
#define S_EXACT 4054.6875

long long N;          // число прямоугольников
int NTh;              // число потоков
int Pin;              // 1 - привязать каждый поток к своему ядру
int NCores;           // число логических процессоров
double dx;
REAL S = 0;           // общая сумма значений функции
CRITICAL_SECTION CS;

DWORD WINAPI ThreadProc(LPVOID param)
{
    int k = (int)(INT_PTR)param;          // номер потока: 0, 1, 2, ...
    long long i;
    REAL x, f, local = 0;

    if (Pin)
        SetThreadAffinityMask(GetCurrentThread(), (DWORD_PTR)1 << (k % NCores));

    for (i = k; i < N; i += NTh) {        // чередование: k, k+NTh, k+2*NTh, ...
        x = A + (i + 0.5) * dx;
        f = 0.05 * x * x * x + 0.3 * x * x - 20.0 * x + 200.0; // f(x) вариант 4
        local = local + f;
    }

    EnterCriticalSection(&CS);
    S = S + local;
    LeaveCriticalSection(&CS);
    return 0;
}

int main(int argc, char *argv[])
{
    HANDLE Thread[64];
    SYSTEM_INFO si;
    clock_t t1, t2;
    double T, Res, Err;
    int k;

    if (argc < 3) {
        printf("usage: integ N NTh [pin] [wait]\n");
        return 1;
    }
    N   = atoll(argv[1]);
    NTh = atoi(argv[2]);
    Pin = (argc > 3) ? atoi(argv[3]) : 0;
    if (NTh < 1 || NTh > 64)
        NTh = 1;

    GetSystemInfo(&si);
    NCores = si.dwNumberOfProcessors;

    if (argc > 4) {                       // пауза, чтобы задать Affinity в Диспетчере задач
        printf("Set affinity and press Enter\n");
        getchar();
    }

    dx = (B - A) / N;
    InitializeCriticalSection(&CS);

    t1 = clock();
    for (k = 0; k < NTh; k++)
        Thread[k] = CreateThread(0, 0, (LPTHREAD_START_ROUTINE)ThreadProc, (LPVOID)(INT_PTR)k, 0, 0);
    WaitForMultipleObjects(NTh, Thread, TRUE, INFINITE);
    t2 = clock();

    DeleteCriticalSection(&CS);
    Res = S * dx;
    T = (double)(t2 - t1) / CLOCKS_PER_SEC;
    Err = (Res - S_EXACT) / S_EXACT * 100;
    printf("%lld,%d,%d,%.10f,%.4f,%e\n", N, NTh, Pin, Res, T, Err);
    return 0;
}



// 
// Раздел 6.7. Тест привязки потоков к ядрам (test.c)
// 
/*
#include <windows.h>
#include <stdio.h>

DWORD WINAPI Work(LPVOID param)
{
    int k = (int)(INT_PTR)param;
    SetThreadAffinityMask(GetCurrentThread(), (DWORD_PTR)1 << k);
    while (1);
    return 0;
}

int main()
{
    DWORD_PTR procMask, sysMask;

    GetProcessAffinityMask(GetCurrentProcess(), &procMask, &sysMask);
    printf("System mask:  %llx\n", (unsigned long long)sysMask);
    printf("Process mask: %llx\n", (unsigned long long)procMask);

    SetProcessAffinityMask(GetCurrentProcess(), 0x0F);      // ядра 0-3
    GetProcessAffinityMask(GetCurrentProcess(), &procMask, &sysMask);
    printf("Process mask after Set: %llx\n", (unsigned long long)procMask);

    CreateThread(0, 0, (LPTHREAD_START_ROUTINE)Work, (LPVOID)(INT_PTR)0, 0, 0);  // ядро 0
    CreateThread(0, 0, (LPTHREAD_START_ROUTINE)Work, (LPVOID)(INT_PTR)2, 0, 0);  // ядро 2
    printf("Threads on cores 0 and 2. Press Enter to exit\n");
    getchar();
    return 0;
}
*/

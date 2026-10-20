
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
// 
// Задание 5.1. Создание параллельного потока с помощью функции Windows API CreateThread()
// 
/*
#include <windows.h>
#include <stdio.h>

void WINAPI Thread2()
{
    printf("Thread Two\n");
    // while(1); // закомментировано для корректного выхода
}

void main()
{
    CreateThread(0, 0, (LPTHREAD_START_ROUTINE)Thread2, 0, 0, 0);
    printf("Main - Thread One\n");
    getchar(); // Ожидание нажатия клавиши для просмотра созданных потоков в Диспетчере задач
}
*/


// 
// Задание 5.3. Параллельные потоки и ситуация гонки данных (Data Race)
// 20 потоков параллельно инкрементируют общую переменную S без синхронизации
// 
/*
#include <windows.h>
#include <stdio.h>
#include <time.h>
#include <stdlib.h>

#define NTh 20

int S  0;
int Ns  10000000;

void WINAPI ThreadProc()
{
    for (int i  0; i < Ns; i++)
        S  S + 1;
}

void main(int argc, char *argv[])
{
    if (argc > 1) {
        Ns  atoi(argv[1]);
    }
    
    HANDLE Thread[NTh];
    clock_t t1, t2;
    float dT;
    
    t1  clock();
    for (int i  0; i < NTh; i++)
        Thread[i]  CreateThread(0, 0, (LPTHREAD_START_ROUTINE)ThreadProc, 0, 0, 0);
        
    WaitForMultipleObjects(NTh, Thread, TRUE, INFINITE);
    t2  clock();
    
    dT  (float)(t2 - t1) / CLOCKS_PER_SEC;
    
    int expected_S  Ns * NTh;
    int error  expected_S - S;
    
    // Вывод: S, Ns, NTh, error, dT
    printf("%d,%d,%d,%d,%5.10f\n", S, Ns, NTh, error, dT);
}
*/


// 
// Задание 5.4. Синхронизация потоков с помощью критической секции (Critical Section)
// Блокировка каждого обращения к общей переменной S (надежно, но медленно)
// 
/*
#include <windows.h>
#include <stdio.h>
#include <time.h>
#include <stdlib.h>

#define NTh 20

int S  0;
int Ns  1000000;
CRITICAL_SECTION CS;

void WINAPI ThreadProc()
{
    for (int i  0; i < Ns; i++)
    {
        EnterCriticalSection(&CS);
        S  S + 1;
        LeaveCriticalSection(&CS);
    }
}

void main(int argc, char *argv[])
{
    if (argc > 1) {
        Ns  atoi(argv[1]);
    }
    
    HANDLE Thread[NTh];
    clock_t t1, t2;
    float dT;
    
    InitializeCriticalSection(&CS);
    
    t1  clock();
    for (int i  0; i < NTh; i++)
        Thread[i]  CreateThread(0, 0, (LPTHREAD_START_ROUTINE)ThreadProc, 0, 0, 0);
        
    WaitForMultipleObjects(NTh, Thread, TRUE, INFINITE);
    t2  clock();
    
    DeleteCriticalSection(&CS);
    
    dT  (float)(t2 - t1) / CLOCKS_PER_SEC;
    
    int expected_S  Ns * NTh;
    int error  expected_S - S;
    
    // Вывод: S, Ns, NTh, error, dT
    printf("%d,%d,%d,%d,%5.10f\n", S, Ns, NTh, error, dT);
}
*/


// 
// Задание 5.5. Оптимизация взаимодействия потоков через локальные переменные
// Потоки накапливают сумму локально, а в критическую секцию входят лишь один раз
// 
/*
#include <windows.h>
#include <stdio.h>
#include <time.h>
#include <stdlib.h>

#define NTh 20

int S  0;
int Ns  1000000;
CRITICAL_SECTION CS;

void WINAPI ThreadProc()
{
    int local_S  0;
    for (int i  0; i < Ns; i++)
    {
        local_S  local_S + 1;
    }
    
    EnterCriticalSection(&CS);
    S  S + local_S;
    LeaveCriticalSection(&CS);
}

void main(int argc, char *argv[])
{
    if (argc > 1) {
        Ns  atoi(argv[1]);
    }
    
    HANDLE Thread[NTh];
    clock_t t1, t2;
    float dT;
    
    InitializeCriticalSection(&CS);
    
    t1  clock();
    for (int i  0; i < NTh; i++)
        Thread[i]  CreateThread(0, 0, (LPTHREAD_START_ROUTINE)ThreadProc, 0, 0, 0);
        
    WaitForMultipleObjects(NTh, Thread, TRUE, INFINITE);
    t2  clock();
    
    DeleteCriticalSection(&CS);
    
    dT  (float)(t2 - t1) / CLOCKS_PER_SEC;
    
    int expected_S  Ns * NTh;
    int error  expected_S - S;
    
    // Вывод: S, Ns, NTh, error, dT
    printf("%d,%d,%d,%d,%5.10f\n", S, Ns, NTh, error, dT);
}
*/


// 
// Задания 6.5, 6.6. Последовательное численное интегрирование методом прямоугольников (Вариант 4)
// Функция: f(x)  0.05*x^3 + 0.3*x^2 - 20*x + 200 на отрезке [-5, 20]
// Точное аналитическое значение интеграла: S  4054.6875
// 
/*
#include <stdio.h>
#include <time.h>
#include <stdlib.h>

float f(float x)
{
    return 0.05f * x * x * x + 0.3f * x * x - 20.0f * x + 200.0f;
}

void main(int argc, char *argv[])
{
    int N  100; // количество отрезков по умолчанию
    if (argc > 1) {
        N  atoi(argv[1]);
    }
    
    float A  -5.0f;
    float B  20.0f;
    float dx  (B - A) / N;
    float S  0.0f;
    
    clock_t t1, t2;
    float dT;
    
    t1  clock();
    for (int i  0; i < N; i++) {
        float x  A + dx * (i + 0.5f);
        S + f(x) * dx;
    }
    t2  clock();
    
    dT  (float)(t2 - t1) / CLOCKS_PER_SEC;
    
    float exact  4054.6875f;
    float err  exact - S;
    float err_pct  (err / exact) * 100.0f;
    if (err_pct < 0) err_pct  -err_pct;
    
    // Вывод: N, Интеграл (S), Погрешность % (err_pct), Время (dT)
    printf("%d,%f,%f,%5.10f\n", N, S, err_pct, dT);
}
*/


// 
// Задание 6.7. Параллельное численное интегрирование с привязкой потоков к ядрам
// Функция: f(x)  0.05*x^3 + 0.3*x^2 - 20*x + 200 на отрезке [-5, 20]
// Метод чередования отрезков и привязка потоков через SetThreadAffinityMask()
// 
/*
#include <windows.h>
#include <stdio.h>
#include <time.h>
#include <stdlib.h>

#define MAX_THREADS 64

float A  -5.0f;
float B  20.0f;
int N  100000;
int num_threads  4;
float global_S  0.0f;
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
    struct ThreadData* data  (struct ThreadData*)lpParam;
    int id  data->id;
    
    // Привязка потока к виртуальному процессору (ядру)
    SetThreadAffinityMask(GetCurrentThread(), (DWORD_PTR)1 << id);
    
    float dx  (B - A) / N;
    float local_S  0.0f;
    
    // Чередование распределения отрезков между потоками
    for (int i  id; i < N; i + num_threads) {
        float x  A + dx * (i + 0.5f);
        local_S + f(x) * dx;
    }
    
    EnterCriticalSection(&CS);
    global_S + local_S;
    LeaveCriticalSection(&CS);
}

void main(int argc, char *argv[])
{
    if (argc > 1) N  atoi(argv[1]);
    if (argc > 2) num_threads  atoi(argv[2]);
    if (num_threads > MAX_THREADS) num_threads  MAX_THREADS;
    
    HANDLE Thread[MAX_THREADS];
    struct ThreadData tData[MAX_THREADS];
    clock_t t1, t2;
    float dT;
    
    InitializeCriticalSection(&CS);
    
    t1  clock();
    for (int i  0; i < num_threads; i++) {
        tData[i].id  i;
        Thread[i]  CreateThread(0, 0, (LPTHREAD_START_ROUTINE)ThreadProc, &tData[i], 0, 0);
    }
        
    WaitForMultipleObjects(num_threads, Thread, TRUE, INFINITE);
    t2  clock();
    
    DeleteCriticalSection(&CS);
    
    dT  (float)(t2 - t1) / CLOCKS_PER_SEC;
    
    float exact  4054.6875f;
    float err  exact - global_S;
    float err_pct  (err / exact) * 100.0f;
    if (err_pct < 0) err_pct  -err_pct;
    
    // Вывод: N, Потоки (num_threads), Интеграл (global_S), Погрешность % (err_pct), Время (dT)
    printf("%d,%d,%f,%f,%5.10f\n", N, num_threads, global_S, err_pct, dT);
}
*/
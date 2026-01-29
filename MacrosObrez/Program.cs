using System;
using System.Runtime.InteropServices;
using System.Threading;
using System.Diagnostics;
using System.Drawing;

class MacrosObrez
{
    // Остальные константы остаются без изменений...
    private const int WH_MOUSE_LL = 14;
    private const int WM_XBUTTONDOWN = 0x020B;
    private const int XBUTTON1 = 0x0001;
    private const int XBUTTON2 = 0x0002;

    private const int MOUSEEVENTF_LEFTDOWN = 0x0002;
    private const int MOUSEEVENTF_LEFTUP = 0x0004;
    private const int MOUSEEVENTF_ABSOLUTE = 0x8000;

    private const int KEYEVENTF_KEYDOWN = 0x0000;
    private const int KEYEVENTF_KEYUP = 0x0002;
    private const int VK_SPACE = 0x20;
    private const int VK_X = 0x58;
    private const int VK_E = 0x45;

    private const int TARGET_X = 820;
    private const int TARGET_Y = 1380;
    private const int DELAY_BETWEEN_CLICKS_MS = 130;
    private const int CLICK_HOLD_DELAY_MS = 15;

    private static IntPtr hookHandle = IntPtr.Zero;
    private static LowLevelMouseProc mouseProc = HookCallback;
    private static bool isMacroRunning = false;
    private static readonly object macroLock = new object();

    static void Main(string[] args)
    {
        Console.WriteLine("MacrosObrez запущен. Нажмите Mouse7 для выполнения макроса.");
        Console.WriteLine("Нажмите Ctrl+C или закройте окно для выхода.");

        // Установка обработчика Ctrl+C для корректного завершения
        Console.CancelKeyPress += (sender, e) =>
        {
            e.Cancel = true;
            Console.WriteLine("\nЗавершение работы...");
            Cleanup();
            Environment.Exit(0);
        };

        hookHandle = SetHook(mouseProc);

        if (hookHandle == IntPtr.Zero)
        {
            Console.WriteLine("ОШИБКА: Не удалось установить хук мыши!");
            Console.WriteLine("Убедитесь, что приложение запущено с правами администратора.");
            return;
        }

        Console.WriteLine("Хук успешно установлен. Ожидание нажатия Mouse7...");

        Application.Run();
        Cleanup();
    }

    // Остальной код остается без изменений...
    [StructLayout(LayoutKind.Sequential)]
    private struct POINT
    {
        public int x;
        public int y;
    }

    [StructLayout(LayoutKind.Sequential)]
    private struct MSLLHOOKSTRUCT
    {
        public POINT pt;
        public uint mouseData;
        public uint flags;
        public uint time;
        public IntPtr dwExtraInfo;
    }

    private delegate IntPtr LowLevelMouseProc(int nCode, IntPtr wParam, IntPtr lParam);

    [DllImport("user32.dll", CharSet = CharSet.Auto, SetLastError = true)]
    private static extern IntPtr SetWindowsHookEx(int idHook, LowLevelMouseProc lpfn, IntPtr hMod, uint dwThreadId);

    [DllImport("user32.dll", CharSet = CharSet.Auto, SetLastError = true)]
    [return: MarshalAs(UnmanagedType.Bool)]
    private static extern bool UnhookWindowsHookEx(IntPtr hhk);

    [DllImport("user32.dll", CharSet = CharSet.Auto, SetLastError = true)]
    private static extern IntPtr CallNextHookEx(IntPtr hhk, int nCode, IntPtr wParam, IntPtr lParam);

    [DllImport("kernel32.dll", CharSet = CharSet.Auto, SetLastError = true)]
    private static extern IntPtr GetModuleHandle(string? lpModuleName);

    [DllImport("user32.dll")]
    private static extern bool GetCursorPos(out POINT lpPoint);

    [DllImport("user32.dll")]
    private static extern bool SetCursorPos(int X, int Y);

    [DllImport("user32.dll")]
    private static extern void mouse_event(int dwFlags, int dx, int dy, int dwData, int dwExtraInfo);

    [DllImport("user32.dll")]
    private static extern void keybd_event(byte bVk, byte bScan, uint dwFlags, int dwExtraInfo);

    // Добавляем импорты для работы с цветом пикселей
    [DllImport("user32.dll")]
    private static extern IntPtr GetDC(IntPtr hWnd);

    [DllImport("user32.dll")]
    private static extern int ReleaseDC(IntPtr hWnd, IntPtr hDC);

    [DllImport("gdi32.dll")]
    private static extern uint GetPixel(IntPtr hdc, int nXPos, int nYPos);

    private static void Cleanup()
    {
        if (hookHandle != IntPtr.Zero)
        {
            UnhookWindowsHookEx(hookHandle);
            hookHandle = IntPtr.Zero;
            Console.WriteLine("Хук удален.");
        }
    }

    private static IntPtr SetHook(LowLevelMouseProc proc)
    {
        using (Process curProcess = Process.GetCurrentProcess())
        using (ProcessModule? curModule = curProcess.MainModule)
        {
            if (curModule != null)
            {
                return SetWindowsHookEx(WH_MOUSE_LL, proc, GetModuleHandle(curModule.ModuleName), 0);
            }
            return IntPtr.Zero;
        }
    }

    private static IntPtr HookCallback(int nCode, IntPtr wParam, IntPtr lParam)
    {
        if (nCode >= 0 && wParam == (IntPtr)WM_XBUTTONDOWN)
        {
            MSLLHOOKSTRUCT hookStruct = Marshal.PtrToStructure<MSLLHOOKSTRUCT>(lParam);

            uint mouseData = hookStruct.mouseData >> 16;
            if (mouseData == XBUTTON2)
            {
                Console.WriteLine("Mouse7 нажата! Выполнение макроса...");
                ExecuteMacro();
            }
            else if (mouseData == XBUTTON1)
            {
                Console.WriteLine("Mouse6 нажата! Выполнение макроса...");
                ExecuteMacro2();
            }
        }

        return CallNextHookEx(hookHandle, nCode, wParam, lParam);
    }

    private static void ExecuteMacro()
    {
        lock (macroLock)
        {
            if (isMacroRunning)
            {
                Console.WriteLine("Макрос уже выполняется, ожидайте завершения...");
                return;
            }
            isMacroRunning = true;
        }

        Thread macroThread = new Thread(() =>
        {
            try
            {
                GetCursorPos(out POINT savedPosition);
                Console.WriteLine($"Позиция сохранена: ({savedPosition.x}, {savedPosition.y})");

                ClickLeftMouse();

                Thread.Sleep(500);

                ClickLeftMouse();

                PressSpace();

                PressX();

                Thread.Sleep(100);

                SetCursorPos(700, 1380);

                Thread.Sleep(100);

                ClickLeftMouse();

                Thread.Sleep(100);

                PressE();

                Thread.Sleep(100);

                ClickLeftMouse();

                Thread.Sleep(100);

                PressE();

                Thread.Sleep(300);

                SetCursorPos(TARGET_X, TARGET_Y);

                Thread.Sleep(DELAY_BETWEEN_CLICKS_MS);

                ClickLeftMouse();

                SetCursorPos(700, 1380);

                Thread.Sleep(DELAY_BETWEEN_CLICKS_MS);

                ClickLeftMouse();

                Thread.Sleep(600);

                ClickLeftMouse();

                Thread.Sleep(600);

                SetCursorPos(TARGET_X, TARGET_Y);
                ClickLeftMouse();

                PressX();

                Thread.Sleep(DELAY_BETWEEN_CLICKS_MS);

                PressSpace();

                SetCursorPos(savedPosition.x, savedPosition.y);
                Console.WriteLine($"Мышь возвращена на ({savedPosition.x}, {savedPosition.y})");

                Console.WriteLine("Макрос выполнен!");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Ошибка при выполнении макроса: {ex.Message}");
            }
            finally
            {
                lock (macroLock)
                {
                    isMacroRunning = false;
                }
            }
        });

        macroThread.Start();
    }

    private static void ExecuteMacro2()
    {
        lock (macroLock)
        {
            if (isMacroRunning)
            {
                Console.WriteLine("Макрос уже выполняется, ожидайте завершения...");
                return;
            }
            isMacroRunning = true;
        }

        Thread macroThread = new Thread(() =>
        {
            try
            {
                ClickLeftMouse();

                Thread.Sleep(500);

                ClickLeftMouse();

                Thread.Sleep(200);

                PressX();

                Thread.Sleep(200);

                ClickLeftMouse();

                Thread.Sleep(500);

                ClickLeftMouse();

                Console.WriteLine("Макрос выполнен!");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Ошибка при выполнении макроса: {ex.Message}");
            }
            finally
            {
                lock (macroLock)
                {
                    isMacroRunning = false;
                }
            }
        });

        macroThread.Start();
    }

    private static void ClickLeftMouse()
    {
        mouse_event(MOUSEEVENTF_LEFTDOWN, 0, 0, 0, 0);
        Thread.Sleep(CLICK_HOLD_DELAY_MS);
        mouse_event(MOUSEEVENTF_LEFTUP, 0, 0, 0, 0);
    }

    private static void PressSpace()
    {
        keybd_event(VK_SPACE, 0, KEYEVENTF_KEYDOWN, 0);
        Thread.Sleep(CLICK_HOLD_DELAY_MS);
        keybd_event(VK_SPACE, 0, KEYEVENTF_KEYUP, 0);
    }

    private static void PressX()
    {
        keybd_event(VK_X, 0, KEYEVENTF_KEYDOWN, 0);
        Thread.Sleep(CLICK_HOLD_DELAY_MS); // Небольшая задержка между нажатием и отпусканием
        keybd_event(VK_X, 0, KEYEVENTF_KEYUP, 0);
    }

    private static void PressE()
    {
        keybd_event(VK_E, 0, KEYEVENTF_KEYDOWN, 0);
        Thread.Sleep(CLICK_HOLD_DELAY_MS); // Небольшая задержка между нажатием и отпусканием
        keybd_event(VK_E, 0, KEYEVENTF_KEYUP, 0);
    }

    private static class Application
    {
        [DllImport("user32.dll")]
        private static extern bool GetMessage(out MSG lpMsg, IntPtr hWnd, uint wMsgFilterMin, uint wMsgFilterMax);

        [DllImport("user32.dll")]
        private static extern bool TranslateMessage([In] ref MSG lpMsg);

        [DllImport("user32.dll")]
        private static extern IntPtr DispatchMessage([In] ref MSG lpmsg);

        [StructLayout(LayoutKind.Sequential)]
        private struct MSG
        {
            public IntPtr hwnd;
            public uint message;
            public IntPtr wParam;
            public IntPtr lParam;
            public uint time;
            public POINT pt;
        }

        public static void Run()
        {
            MSG msg;
            while (GetMessage(out msg, IntPtr.Zero, 0, 0))
            {
                TranslateMessage(ref msg);
                DispatchMessage(ref msg);
            }
        }
    }
}

using System;
using System.Runtime.InteropServices;
using System.Threading;
using System.Diagnostics;

class MacrosObrez
{
    // Константы для мыши и клавиатуры
    private const int WH_MOUSE_LL = 14;
    private const int WM_XBUTTONDOWN = 0x020B;
    private const int XBUTTON2 = 0x0002; // Mouse7 (обычно вторая дополнительная кнопка)
    
    private const int MOUSEEVENTF_LEFTDOWN = 0x0002;
    private const int MOUSEEVENTF_LEFTUP = 0x0004;
    private const int MOUSEEVENTF_ABSOLUTE = 0x8000;
    
    private const int KEYEVENTF_KEYDOWN = 0x0000;
    private const int KEYEVENTF_KEYUP = 0x0002;
    private const int VK_SPACE = 0x20;
    
    // Константы для настройки макроса
    private const int TARGET_X = 200;
    private const int TARGET_Y = 200;
    private const int DELAY_BETWEEN_CLICKS_MS = 150;
    private const int CLICK_HOLD_DELAY_MS = 10;
    
    private static IntPtr hookHandle = IntPtr.Zero;
    private static LowLevelMouseProc mouseProc = HookCallback;
    private static bool isMacroRunning = false;
    private static readonly object macroLock = new object();
    
    // Структуры для работы с хуками
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
    
    // Делегат для обработки хука мыши
    private delegate IntPtr LowLevelMouseProc(int nCode, IntPtr wParam, IntPtr lParam);
    
    // Импорт функций Windows API
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
    
    static void Main(string[] args)
    {
        Console.WriteLine("MacrosObrez запущен. Нажмите Mouse7 для выполнения макроса.");
        Console.WriteLine("Нажмите Ctrl+C для выхода.");
        
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
        
        // Ожидание завершения
        Application.Run();
        
        Cleanup();
    }
    
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
            
            // Проверяем, что нажата именно Mouse7 (XBUTTON2)
            uint mouseData = hookStruct.mouseData >> 16;
            if (mouseData == XBUTTON2)
            {
                Console.WriteLine("Mouse7 нажата! Выполнение макроса...");
                ExecuteMacro();
            }
        }
        
        return CallNextHookEx(hookHandle, nCode, wParam, lParam);
    }
    
    private static void ExecuteMacro()
    {
        // Проверяем, не выполняется ли уже макрос
        lock (macroLock)
        {
            if (isMacroRunning)
            {
                Console.WriteLine("Макрос уже выполняется, ожидайте завершения...");
                return;
            }
            isMacroRunning = true;
        }
        
        // Запускаем макрос в отдельном потоке, чтобы не блокировать хук
        Thread macroThread = new Thread(() =>
        {
            try
            {
                // 1. Сохранить позицию мыши
                GetCursorPos(out POINT savedPosition);
                Console.WriteLine($"Позиция сохранена: ({savedPosition.x}, {savedPosition.y})");
                
                // 2. Нажать левую кнопку мыши
                ClickLeftMouse();
                Console.WriteLine("Клик 1");
                
                // 3. Подождать 0.15 секунды
                Thread.Sleep(DELAY_BETWEEN_CLICKS_MS);
                
                // 4. Нажать левую кнопку мыши
                ClickLeftMouse();
                Console.WriteLine("Клик 2");
                
                // 5. Нажать пробел на клавиатуре
                PressSpace();
                Console.WriteLine("Пробел нажат");
                
                // 6. Передвинуть мышку на 200, 200
                SetCursorPos(TARGET_X, TARGET_Y);
                Console.WriteLine($"Мышь перемещена на ({TARGET_X}, {TARGET_Y})");
                
                // 7. Нажать левую кнопку мыши
                ClickLeftMouse();
                Console.WriteLine("Клик 3");
                
                // 8. Вернуться на сохраненную позицию
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
    
    private static void ClickLeftMouse()
    {
        mouse_event(MOUSEEVENTF_LEFTDOWN, 0, 0, 0, 0);
        Thread.Sleep(CLICK_HOLD_DELAY_MS); // Небольшая задержка между нажатием и отпусканием
        mouse_event(MOUSEEVENTF_LEFTUP, 0, 0, 0, 0);
    }
    
    private static void PressSpace()
    {
        keybd_event(VK_SPACE, 0, KEYEVENTF_KEYDOWN, 0);
        Thread.Sleep(CLICK_HOLD_DELAY_MS); // Небольшая задержка между нажатием и отпусканием
        keybd_event(VK_SPACE, 0, KEYEVENTF_KEYUP, 0);
    }
    
    // Простая реализация цикла сообщений
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

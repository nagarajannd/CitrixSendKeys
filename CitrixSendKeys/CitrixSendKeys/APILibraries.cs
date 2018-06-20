using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Runtime.InteropServices;
using System.Diagnostics;

namespace CitrixSendKeys
{
    public class APILibraries
    {
        #region API Methods
        [DllImport("user32.dll", SetLastError = true)]
        protected static extern IntPtr FindWindowEx(IntPtr hwndParent, IntPtr hwndChildAfter, string lpszClass, string lpszWindow);
        [DllImport("user32.dll", SetLastError = true, CharSet = CharSet.Auto)]
        protected static extern int GetClassName(IntPtr hWnd, StringBuilder lpClassName, int nMaxCount);
        [DllImport("user32.dll", CharSet = CharSet.Auto, SetLastError = true)]
        protected static extern int GetWindowText(IntPtr hWnd, StringBuilder lpString, int nMaxCount);

        [DllImport("user32.dll")]
        protected static extern int MapVirtualKey(int uCode, int uMapType);
        [DllImport("user32.dll", CharSet = CharSet.Unicode)]
        protected static extern int VkKeyScan(char ch);
        [DllImport("user32.dll")]
        protected static extern bool GetKeyState(int nVirtKey);

        [DllImport("user32.dll")]
        [return: MarshalAs(UnmanagedType.Bool)]
        protected static extern bool SetForegroundWindow(IntPtr hWnd);
        [DllImport("user32.dll")]
        protected static extern IntPtr GetForegroundWindow();
        [DllImport("user32.dll", EntryPoint = "ShowWindow", SetLastError = true)]
        protected static extern bool ShowWindow(IntPtr hWnd, int nCmdShow);

        [DllImport("user32.dll", SetLastError = true)]
        [return: MarshalAs(UnmanagedType.Bool)]
        protected static extern bool GetWindowPlacement(IntPtr hWnd, ref WINDOWPLACEMENT lpwndpl);
        [DllImport("user32.dll", SetLastError = true)]
        protected static extern bool GetWindowRect(IntPtr hwnd, out RECT lpRect);

        [DllImport("user32.dll")]
        protected static extern void keybd_event(byte bVk, byte bScan, uint dwFlags, UIntPtr dwExtraInfo);
        [DllImport("user32.dll")]
        protected static extern void mouse_event(uint dwFlags, int dx, int dy, uint dwData, UIntPtr dwExtraInfo);

        [DllImport("user32.dll")]
        [return: MarshalAs(UnmanagedType.Bool)]
        protected static extern bool SetCursorPos(int x, int y);

        [DllImport("user32.dll")]
        [return: MarshalAs(UnmanagedType.Bool)]
        protected static extern bool SetWindowPos(IntPtr hWnd, IntPtr hWndInsertAfter, int X, int Y, int cx, int cy, uint uFlags);

        [DllImport("kernel32.dll")]
        protected static extern IntPtr OpenThread(ThreadAccess dwDesiredAccess, bool bInheritHandle, uint dwThreadId);
        [DllImport("kernel32.dll")]
        protected static extern uint SuspendThread(IntPtr hThread);
        [DllImport("kernel32.dll")]
        protected static extern int ResumeThread(IntPtr hThread);
        [DllImport("kernel32", CharSet = CharSet.Auto, SetLastError = true)]
        protected static extern bool CloseHandle(IntPtr handle);

        [DllImport("User32.dll", SetLastError = true)]
        [return: MarshalAs(UnmanagedType.Bool)]
        protected static extern bool IsClipboardFormatAvailable(uint format);
        [DllImport("User32.dll", SetLastError = true)]
        protected static extern IntPtr GetClipboardData(uint uFormat);
        [DllImport("User32.dll", SetLastError = true)]
        [return: MarshalAs(UnmanagedType.Bool)]
        protected static extern bool OpenClipboard(IntPtr hWndNewOwner);
        [DllImport("User32.dll", SetLastError = true)]
        [return: MarshalAs(UnmanagedType.Bool)]
        protected static extern bool CloseClipboard();
        [DllImport("user32.dll")]
        [return: MarshalAs(UnmanagedType.Bool)]
        protected static extern bool EmptyClipboard();

        [DllImport("Kernel32.dll", SetLastError = true)]
        protected static extern IntPtr GlobalLock(IntPtr hMem);
        [DllImport("Kernel32.dll", SetLastError = true)]
        [return: MarshalAs(UnmanagedType.Bool)]
        protected static extern bool GlobalUnlock(IntPtr hMem);
        [DllImport("Kernel32.dll", SetLastError = true)]
        protected static extern int GlobalSize(IntPtr hMem);
        #endregion

        #region Enumurables
        [Serializable]
        [StructLayout(LayoutKind.Sequential)]
        protected struct WINDOWPLACEMENT
        {
            public int Length;
            public int Flags;
            public ShowWindowCommands ShowCmd;
            public POINT MinPosition;
            public POINT MaxPosition;
            public RECT NormalPosition;
        }
        protected enum ThreadAccess : int
        {
            TERMINATE = (0x0001),
            SUSPEND_RESUME = (0x0002),
            GET_CONTEXT = (0x0008),
            SET_CONTEXT = (0x0010),
            SET_INFORMATION = (0x0020),
            QUERY_INFORMATION = (0x0040),
            SET_THREAD_TOKEN = (0x0080),
            IMPERSONATE = (0x0100),
            DIRECT_IMPERSONATION = (0x0200)
        }
        protected struct RECT
        {
            public int Left;
            public int Top;
            public int Right;
            public int Bottom;
        }
        protected struct POINT
        {
            public int x;
            public int y;
        }
        protected struct FIXED
        {
            public short fract;
            public short value;
        }
        protected enum ShowWindowCommands : int
        {
            Hide = 0,
            Normal = 1,
            ShowMinimized = 2,
            Maximize = 3,
            ShowMaximized = 3,
            ShowNoActivate = 4,
            Show = 5,
            Minimize = 6,
            ShowMinNoActive = 7,
            ShowNA = 8,
            Restore = 9,
            ShowDefault = 10,
            ForceMinimize = 11
        }
        protected enum KeyBoardEventEnums : int
        {
            KEYEVENTF_EXTENDEDKEY = 0x001,
            KEYEVENTF_KEYUP = 0x002
        }
        protected enum MouseEventEnums : int
        {
            MOUSEEVENTF_LEFTDOWN = 0x002,
            MOUSEEVENTF_LEFTUP = 0x004,
            MOUSEEVENTF_RIGHTDOWN = 0x008,
            MOUSEEVENTF_RIGHTUP = 0x0010
        }
        protected enum VirtualKeyCodes : int
        {
            keyBackspace = 0x008, keyTab = 0x009, keyReturn = 0x00D, keyEnter = 0x00D,
            keyShift = 0x0010, keyControl = 0x0011, keyAlt = 0x0012, keyPause = 0x0013,
            keyEscape = 0x001B, keySpace = 0x0020, keyEnd = 0x0023, keyHome = 0x0024,
            keyLeft = 0x0025, KeyUp = 0x0026, keyRight = 0x0027, KeyDown = 0x0028,
            keyInsert = 0x002D, keyDelete = 0x002E, keyF1 = 0x0070, keyF2 = 0x0071,
            keyF3 = 0x0072, keyF4 = 0x0073, keyF5 = 0x0074, keyF6 = 0x0075,
            keyF7 = 0x0076, keyF8 = 0x0077, keyF9 = 0x0078, keyF10 = 0x0079,
            keyF11 = 0x007A, keyF12 = 0x007B, KeyPageUp = 0x0021, KeyPageDown = 0x0022,
            KeyPrintScreen = 0x002C, KeyLeftWin = 0x005B, KeyBreak = 0x0022,
            keyNumLock = 0x0090, keyScrollLock = 0x0091, keyCapsLock = 0x0014,
        }
        protected enum eLockKey : int
        {
            CapsLock = VirtualKeyCodes.keyCapsLock,
            NumLock = VirtualKeyCodes.keyNumLock,
            ScrollLock = VirtualKeyCodes.keyScrollLock
        }

        protected static readonly IntPtr HWND_TOPMOST = new IntPtr(-1);
        protected const UInt32 SWP_NOSIZE = 0x0001;
        protected const UInt32 SWP_NOMOVE = 0x0002;
        protected const UInt32 TOPMOST_FLAGS = SWP_NOMOVE | SWP_NOSIZE;
        protected const int WM_WINDOWPOSCHANGING = 0x0046;
        protected const uint CF_UNICODETEXT = 13U;
        #endregion

        #region Public Methods
        #endregion
    }
}

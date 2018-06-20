using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading;

namespace CitrixSendKeys
{
    public class SimulateKeyPress : APILibraries
    {
        int zlKeyPressDelay = 20;
        public SimulateKeyPress()
        {
            if (GetLockStatus((int)VirtualKeyCodes.keyCapsLock))
            {
                PressKeyVK((int)VirtualKeyCodes.keyCapsLock, true, false);
                PressKeyVK((int)VirtualKeyCodes.keyCapsLock, false, true);
            }
            if (GetLockStatus((int)VirtualKeyCodes.keyScrollLock))
            {
                PressKeyVK((int)VirtualKeyCodes.keyScrollLock, true, false);
                PressKeyVK((int)VirtualKeyCodes.keyScrollLock, false, true);
            }
            if (!GetLockStatus((int)VirtualKeyCodes.keyNumLock))
            {
                PressKeyVK((int)VirtualKeyCodes.keyNumLock, true, false);
                PressKeyVK((int)VirtualKeyCodes.keyNumLock, false, true);
            }
        }

        protected void InsertKeystroke(string TypeInText, int Delay)
        {
            List<string> colText = new List<string>();

            colText = SplitTypeinText(TypeInText);
            zlKeyPressDelay = Delay / colText.Count;

            foreach (string sText in colText)
            {
                if (sText.Length > 1)
                {
                    if (sText.Contains("[SHIFT") || sText.Contains("[CTRL") || sText.Contains("[ALT") || sText.Contains("[WIN"))
                    {
                        if (sText.Contains("UP]"))
                            PressKeyVK(getVirtualKey(sText), false, true);
                        if (sText.Contains("DOWN]"))
                            PressKeyVK(getVirtualKey(sText), true, false);
                    }
                    else
                    {
                        PressKeyVK(getVirtualKey(sText), true, false);
                        PressKeyVK(getVirtualKey(sText), false, true);
                    }
                }
                else
                    SendKey(sText, false, false);
                Thread.Sleep(zlKeyPressDelay);
            }
        }
        private List<string> SplitTypeinText(string sText)
        {
            List<string> opColls = new List<string>();
            string sKey;
            int lLen, i, lFrom = 0, lTo = 0;
            bool spKey = false, spKeyEnd = false;

            lLen = sText.Length;
            for (i = 0; i < lLen; i++)
            {
                sKey = sText.Substring(i, 1);
                if (sKey.Equals("[") && !spKey)
                {
                    spKey = true;
                    lFrom = i;
                }
                if (sKey.Equals("]") && !spKeyEnd)
                {
                    spKeyEnd = true;
                    lTo = i - lFrom + 1;
                }
                if (!spKey)
                    opColls.Add(sKey);
                else
                {
                    if (spKeyEnd)
                    {
                        sKey = sText.Substring(lFrom, lTo);
                        opColls.Add(sKey);
                        spKey = spKeyEnd = false;
                    }
                }
            }
            return opColls;
        }

        private void PressKeyVK(int eKeys, bool bHoldKeydown, bool bRelease)
        {
            int lScan, lExtended;

            lScan = MapVirtualKey(eKeys, 1);
            lExtended = 0;
            if (lScan == 0)
                lExtended = (int)KeyBoardEventEnums.KEYEVENTF_EXTENDEDKEY;
            lScan = MapVirtualKey(eKeys, 0);

            if (!bRelease)
                keybd_event((byte)eKeys, (byte)lScan, (uint)lExtended, UIntPtr.Zero);
            if (!bHoldKeydown)
                keybd_event((byte)eKeys, (byte)lScan, (uint)KeyBoardEventEnums.KEYEVENTF_KEYUP | (uint)lExtended, UIntPtr.Zero);
        }
        private void SendKey(string sKey, bool bHoldKeydown, bool bRelease)
        {
            int lScan, lExtended, lVK;
            bool bShift, bCtrl, bAlt;
            char sChar;

            sChar = sKey.ToCharArray()[0];
            lVK = (int)VkKeyScan(sChar);

            if (lVK > 0)
            {
                lScan = MapVirtualKey(lVK, 1);
                lExtended = 0;
                if (lScan == 0)
                    lExtended = (int)KeyBoardEventEnums.KEYEVENTF_EXTENDEDKEY;
                lScan = MapVirtualKey(lVK, 0);

                bShift = (lVK & 0x100) > 0;
                bCtrl = (lVK & 0x200) > 0;
                bAlt = (lVK & 0x400) > 0;
                lVK = (lVK & 0xFF);

                if (!bRelease)
                {
                    if (bShift)
                        keybd_event((byte)VirtualKeyCodes.keyShift, (byte)0, (uint)0, UIntPtr.Zero);
                    if (bCtrl)
                        keybd_event((byte)VirtualKeyCodes.keyControl, (byte)0, (uint)0, UIntPtr.Zero);
                    if (bAlt)
                        keybd_event((byte)VirtualKeyCodes.keyAlt, (byte)0, (uint)0, UIntPtr.Zero);
                    keybd_event((byte)lVK, (byte)lScan, (uint)lExtended, UIntPtr.Zero);
                }
                if (!bHoldKeydown)
                {
                    keybd_event((byte)lVK, (byte)lScan, (uint)KeyBoardEventEnums.KEYEVENTF_KEYUP | (uint)lExtended, UIntPtr.Zero);
                    if (bShift)
                        keybd_event((byte)VirtualKeyCodes.keyShift, (byte)0, (uint)KeyBoardEventEnums.KEYEVENTF_KEYUP, UIntPtr.Zero);
                    if (bCtrl)
                        keybd_event((byte)VirtualKeyCodes.keyControl, (byte)0, (uint)KeyBoardEventEnums.KEYEVENTF_KEYUP, UIntPtr.Zero);
                    if (bAlt)
                        keybd_event((byte)VirtualKeyCodes.keyAlt, (byte)0, (uint)KeyBoardEventEnums.KEYEVENTF_KEYUP, UIntPtr.Zero);
                }
            }

        }

        private bool GetLockStatus(int LockKey)
        {
            return GetKeyState(LockKey);
        }
        private int getVirtualKey(string sText)
        {
            int opKey = 0;
            if (sText.Contains("[SHIFT")) opKey = (int)VirtualKeyCodes.keyShift;
            if (sText.Contains("[CTRL")) opKey = (int)VirtualKeyCodes.keyControl;
            if (sText.Contains("[ALT")) opKey = (int)VirtualKeyCodes.keyAlt;
            if (sText.Contains("[UP")) opKey = (int)VirtualKeyCodes.KeyUp;
            if (sText.Contains("[DOWN")) opKey = (int)VirtualKeyCodes.KeyDown;
            if (sText.Contains("[RIGHT")) opKey = (int)VirtualKeyCodes.keyRight;
            if (sText.Contains("[TAB")) opKey = (int)VirtualKeyCodes.keyTab;
            if (sText.Contains("[F1")) opKey = (int)VirtualKeyCodes.keyF1;
            if (sText.Contains("[F2")) opKey = (int)VirtualKeyCodes.keyF2;
            if (sText.Contains("[F3")) opKey = (int)VirtualKeyCodes.keyF3;
            if (sText.Contains("[F4")) opKey = (int)VirtualKeyCodes.keyF4;
            if (sText.Contains("[F5")) opKey = (int)VirtualKeyCodes.keyF5;
            if (sText.Contains("[F6")) opKey = (int)VirtualKeyCodes.keyF6;
            if (sText.Contains("[F7")) opKey = (int)VirtualKeyCodes.keyF7;
            if (sText.Contains("[F8")) opKey = (int)VirtualKeyCodes.keyF8;
            if (sText.Contains("[F9")) opKey = (int)VirtualKeyCodes.keyF9;
            if (sText.Contains("[F10")) opKey = (int)VirtualKeyCodes.keyF10;
            if (sText.Contains("[F11")) opKey = (int)VirtualKeyCodes.keyF11;
            if (sText.Contains("[F12")) opKey = (int)VirtualKeyCodes.keyF12;
            if (sText.Contains("[BACKSPACE")) opKey = (int)VirtualKeyCodes.keyBackspace;
            if (sText.Contains("[CAPS")) opKey = (int)VirtualKeyCodes.keyCapsLock;
            if (sText.Contains("[HOME")) opKey = (int)VirtualKeyCodes.keyHome;
            if (sText.Contains("[END")) opKey = (int)VirtualKeyCodes.keyEnd;
            if (sText.Contains("[INSERT")) opKey = (int)VirtualKeyCodes.keyInsert;
            if (sText.Contains("[DELETE")) opKey = (int)VirtualKeyCodes.keyDelete;
            if (sText.Contains("[ESC")) opKey = (int)VirtualKeyCodes.keyEscape;
            if (sText.Contains("[ENTER")) opKey = (int)VirtualKeyCodes.keyReturn;
            if (sText.Contains("[NUM LOCK")) opKey = (int)VirtualKeyCodes.keyNumLock;
            if (sText.Contains("[SCROLL")) opKey = (int)VirtualKeyCodes.keyScrollLock;
            if (sText.Contains("[ENTER NUM")) opKey = (int)VirtualKeyCodes.keyEnter;
            if (sText.Contains("[PAGE UP")) opKey = (int)VirtualKeyCodes.KeyPageUp;
            if (sText.Contains("[PAGE DOWN")) opKey = (int)VirtualKeyCodes.KeyPageDown;
            if (sText.Contains("[WIN")) opKey = (int)VirtualKeyCodes.KeyLeftWin;
            return opKey;
        }
    }
}

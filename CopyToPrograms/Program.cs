using System;
using System.Windows;
using Microsoft.Win32;
using Microsoft.VisualBasic;
using System.Security.Principal;
using System.IO;
using System.Security.AccessControl;

using WshShell = IWshRuntimeLibrary.WshShell;
using IWshShortcut = IWshRuntimeLibrary.IWshShortcut;

namespace CopyToPrograms
{
    public class Program
    {
        private const string searchableKeyName = "Pin to Searchable";
        private const string commandKeyName = "command";

        public static void Main(string[] args)
        {
            if (args.Length == 0)
            {
                InstallMe();
                return;
            }
            else if (args.Length == 1)
            {
                string target = args[0];
                Pin(target);
            }
            else
            {
                Console.WriteLine("Too many args");
                Console.ReadLine();
            }
        }

        public static string GetShortcutName(string defaultName)
        {
            return Interaction.InputBox("Enter shortcut name","Pin to Searchable",defaultName);
        }

        public static void Pin(string target)
        {
            string targetWithoutPath = GetShortcutName(StripSlash(target));

            string shortcutPath = @"C:\ProgramData\Microsoft\Windows\Start Menu\Programs\" + targetWithoutPath + ".lnk";

            WshShell shell = new WshShell();
            IWshShortcut shortcut = shell.CreateShortcut(shortcutPath);
            shortcut.TargetPath = target;
            //shortcut.
            shortcut.Save();
        }

        public static string StripSlash(string m_string)
        {
            int lastSlash = m_string.LastIndexOfAny(new char[] { '\\', '/' });
            // strip everything before and including the last slash
            m_string = m_string.Substring(lastSlash + 1);

            return m_string;
        }

        private static void InstallMe()
        {
            if (!IsAdministrator())
            {
                Console.WriteLine("Run with admin to install");
                Console.ReadLine();
                return;
            }
            RegistryKey shellKey = 
                Registry.ClassesRoot.
                OpenSubKey("*",true).
                OpenSubKey("shell",true);
            RegistryKey searchableKey = shellKey.OpenSubKey(searchableKeyName,true);
            if (searchableKey==null)
            {
                // is not installed yet, install it
                Console.WriteLine("Installing...");
                searchableKey = shellKey.CreateSubKey(searchableKeyName);
                string exePath = CopyProgramToProgramFiles();
                InstallSetupRegistry(searchableKey,exePath);
            }
            else
            {
                Console.WriteLine("already installed");
                Console.ReadLine();
            }
        }

        public static bool IsAdministrator()
        {
            using (WindowsIdentity identity = WindowsIdentity.GetCurrent())
            {
                WindowsPrincipal principal = new WindowsPrincipal(identity);
                return principal.IsInRole(WindowsBuiltInRole.Administrator);
            }
        }

        private static string CopyProgramToProgramFiles()
        {
            string programFilePath = Environment.GetFolderPath(Environment.SpecialFolder.ProgramFiles);
            string searchablePath = programFilePath + "\\PinToSearchable";
            Directory.CreateDirectory(searchablePath);
            string exePath = searchablePath + "\\PinToSearchable.exe";
            File.Copy(NameOfThisExe(), exePath);
            return exePath;
        }

        private static string NameOfThisExe()
        {
            return System.Reflection.Assembly.GetEntryAssembly().Location;
        }

        private static void InstallSetupRegistry(RegistryKey searchableKey, string exePath)
        {
            var commandKey = searchableKey.CreateSubKey(commandKeyName);
            string cmd = exePath + " \"%1\"";
            commandKey.SetValue("", cmd, RegistryValueKind.String);
        }

    }
}

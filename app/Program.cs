using System;
using System.IO;
using System.Diagnostics;
using System.Text.RegularExpressions;

namespace PortableR
{
    class Program
    {
        static void Main(string[] args)
        {
            // init vars
            string[] arguments = Environment.GetCommandLineArgs();
            string portablerExe = arguments[0]; // C:\Projects\portabler\app\bin\PortableR.exe
            string portablerFolder = Path.GetDirectoryName(portablerExe) + "\\"; // C:\Projects\portabler\app\bin\
            string action = arguments.Length > 1 ? arguments[1] : "";

            // create log file
            StreamWriter log = new StreamWriter(portablerFolder + "debug.log");
            log.WriteLine("\n   ___           _        _     _        __  ");
            log.WriteLine("  / _ \\___  _ __| |_ __ _| |__ | | ___  /__\\ ");
            log.WriteLine(" / /_)/ _ \\| '__| __/ _` | '_ \\| |/ _ \\/ \\// ");
            log.WriteLine("/ ___/ (_) | |  | || (_| | |_) | |  __/ _  \\ ");
            log.WriteLine("\\/    \\___/|_|   \\__\\__,_|_.__/|_|\\___\\/ \\_/ ");

            // log parameters
            log.WriteLine("\n\n=== Variables ===");
            log.WriteLine("arguments       : " + String.Join(", ", arguments));
            log.WriteLine("portablerExe    : " + portablerExe);
            log.WriteLine("portablerFolder : " + portablerFolder);

            if (action.Length > 0)
            {
                log.WriteLine("action          : " + action);
            }

            log.Flush();

            if (action == "install")
            {
                // set contextual menu to registry
                Microsoft.Win32.RegistryKey key = Microsoft.Win32.Registry.ClassesRoot;
                try
                {
                    key = key.CreateSubKey("exefile\\shell\\PortableR\\command");
                    key.SetValue("", arguments[0] + " create \"%0\"");
                }
                catch (Exception ex)
                {
                    log.WriteLine("error during wrinting to reg : " + ex);
                }
                finally
                {
                    key.Close();
                }
            }
            else if (action == "create")
            {
                string appExe = arguments[2]; // "C:\\My Apps\\Text Reader\\tr2.exe";            
                string appName = Path.GetFileName(appExe); // "tr2.exe"               
                string appDir = Path.GetDirectoryName(appExe); // "C:\\My Apps\\Text Reader"      
                log.WriteLine("appExe          : " + appExe);
                log.WriteLine("appName         : " + appName);
                log.WriteLine("appDir          : " + appDir);
                log.Flush();

                // set portable app temp & final name
                string portableAppTemp = "portableApp.exe";
                string portableAppName = appDir.Substring(appDir.LastIndexOf('\\') + 1).Replace(" ", "."); // "Text.Reader"               
                var versionInfo = FileVersionInfo.GetVersionInfo(appExe);
                string version = (versionInfo.ProductVersion != null && versionInfo.ProductVersion.Length > 0) ? versionInfo.ProductVersion : "0"; // 9.5.3
                version = Regex.Replace(version, @"\D", "");
                string fileVersion = (versionInfo.FileVersion != null && versionInfo.FileVersion.Length > 0) ? versionInfo.FileVersion : "0"; // 9.5.3
                fileVersion = Regex.Replace(fileVersion, @"\D", "");
                string finalVersion = (int.Parse(version) > int.Parse(fileVersion)) ? version : fileVersion;
                portableAppName += ("_" + (finalVersion + "000").Substring(0, 3) + ".exe");  // "Text.Reader_200.exe"
                appDir += "\\";
                log.WriteLine("version         : " + version);
                log.WriteLine("fileVersion     : " + fileVersion);
                log.WriteLine("finalVersion    : " + finalVersion);
                log.WriteLine("portableAppTemp : " + portableAppTemp);
                log.WriteLine("portableAppName : " + portableAppName);
                log.Flush();

                // create sfx config
                StreamWriter configSFX = new StreamWriter(portablerFolder + "ConfigSFX.txt");
                configSFX.WriteLine("TempMode");
                configSFX.WriteLine("Silent=2");
                configSFX.WriteLine("Overwrite=2");
                configSFX.WriteLine("Title=Portable App by PortableR");
                configSFX.WriteLine("Setup=" + appName);
                configSFX.Flush();
                configSFX.Close();

                string cmd, parameters;

                // create sfx               
                try
                {
                    cmd = portablerFolder + "third-party\\" + "rar.exe";
                    parameters = " " + "a -ep -ep1 -r -sfx\"" + portablerFolder + "Default.sfx\" -z\"" + portablerFolder + "ConfigSFX.txt\" \"" + appDir + portableAppTemp + "\" \"" + appDir + "*\"";
                    log.WriteLine("\n=== SFX Creation ===");
                    log.WriteLine("starting cmd    : " + cmd);
                    log.WriteLine("with params     :" + parameters);
                    Process sfxProcess = Process.Start(cmd, parameters);
                    sfxProcess.WaitForExit();
                }
                catch (Exception ex)
                {
                    log.WriteLine("error while producing sfx : " + ex.Message);
                }
                finally
                {
                    log.Flush();
                }


                // extract icon                
                try
                {
                    cmd = portablerFolder + "third-party\\" + "ResourceHacker.exe";
                    parameters = " " + "-extract" + " " + "\"" + appExe + "\"" + ", MyProgIcons.rc, icongroup,,";
                    log.WriteLine("\n=== Icon Extract ===");
                    log.WriteLine("starting cmd    : " + cmd);
                    log.WriteLine("with params     :" + parameters);
                    Process extractProcess = Process.Start(cmd, parameters);
                    extractProcess.WaitForExit();

                }
                catch (Exception ex)
                {
                    log.WriteLine("error while extracting icon : " + ex.Message);
                }
                finally
                {
                    log.Flush();
                }

                // inject icon
                try
                {
                    log.WriteLine("\n=== Icon Injection ===");
                    long iconFileSize = new FileInfo(appDir + "Icon_1.ico").Length;
                    log.WriteLine("icon file size  : " + iconFileSize);
                    if (iconFileSize > 1)
                    {
                        cmd = portablerFolder + "third-party\\" + "ResourceHacker.exe";
                        parameters = " " + "-addoverwrite" + " " + "\"" + appDir + portableAppTemp + "\"" + "," + portableAppName + ",Icon_1.ico,ICONGROUP,MAINICON,0";
                        log.WriteLine("starting cmd    : " + cmd);
                        log.WriteLine("with params     :" + parameters);
                        Process injectProcess = Process.Start(cmd, parameters);
                        injectProcess.WaitForExit();
                    }
                    else
                    {
                        log.WriteLine("no icon inject  : icon file size around zero bytes");
                        File.Move(appDir + portableAppTemp, appDir + portableAppName);
                    }
                }
                catch (Exception ex)
                {
                    log.WriteLine("error while injecting icon : " + ex.Message);
                }
                finally
                {
                    log.Flush();
                }

                // temporary file cleaning
                try
                {
                    log.WriteLine("\n=== Temp File Cleaning ===");
                    log.WriteLine("deleting file   : " + portableAppTemp);
                    File.Delete(appDir + portableAppTemp);
                    log.WriteLine("deleting file   : " + "MyProgIcons.rc");
                    File.Delete(appDir + "MyProgIcons.rc");
                    string[] icons = Directory.GetFiles(appDir, "*.ico");
                    foreach (string icon in icons)
                    {
                        log.WriteLine("deleting file   : " + icon);
                        File.Delete(icon);
                    }

                }
                catch (Exception ex)
                {
                    log.WriteLine("error while cleaning temp files : " + ex.Message);
                }

            }
            else if (action == "")
            {
                log.WriteLine("no action given");
            }
            else
            {
                log.WriteLine("non handled action");
            }

            // flush and close log
            log.Flush();
            log.Close();
        }
    }
}

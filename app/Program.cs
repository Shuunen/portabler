using System;
using System.IO;
using System.Diagnostics;
using System.Text.RegularExpressions;

namespace PortableR
{
    class Program
    {
        const string logFile = "debug.log";
        static StreamWriter logInstance;
        static bool hadError = false;

        static void Main(string[] args)
        {
            // init
            string[] arguments = Environment.GetCommandLineArgs();
            string portablerExe = arguments[0]; // C:\Projects\portabler\app\bin\PortableR.exe
            string portablerFolder = Path.GetDirectoryName(portablerExe) + "\\"; // C:\Projects\portabler\app\bin\
            string action = arguments.Length > 1 ? arguments[1] : "";
            logInstance = new StreamWriter(portablerFolder + logFile);
            logInstance.AutoFlush = true;

            // splash
            Log("\n", "draw");
            Log("   ___           _        _     _        __  ", "draw");
            Log("  / _ \\___  _ __| |_ __ _| |__ | | ___  /__\\ ", "draw");
            Log(" / /_)/ _ \\| '__| __/ _` | '_ \\| |/ _ \\/ \\// ", "draw");
            Log("/ ___/ (_) | |  | || (_| | |_) | |  __/ _  \\ ", "draw");
            Log("\\/    \\___/|_|   \\__\\__,_|_.__/|_|\\___\\/ \\_/ ", "draw");

            // log parameters
            Log("Variables", "title");
            Log("arguments : " + String.Join(", ", arguments));
            Log("portablerExe : " + portablerExe);
            Log("portablerFolder : " + portablerFolder);

            if (action.Length > 0)
            {
                Log("action : " + action);
            }

			switch (action) {
				case "install":
					Microsoft.Win32.RegistryKey key = Microsoft.Win32.Registry.ClassesRoot;
					try {
						key = key.CreateSubKey("exefile\\shell\\PortableR\\command");
						key.SetValue("", arguments[0] + " create \"%0\"");
					} catch (Exception ex) {
						Log("error during wrinting to reg : " + ex, "error");
					} finally {
						key.Close();
					}
					break;
				case "create":
					string appExe = arguments[2];
					string appName = Path.GetFileName(appExe);
					string appDir = Path.GetDirectoryName(appExe);
					Log("appExe : " + appExe);
					Log("appName : " + appName);
					Log("appDir : " + appDir);
					const string portableAppTemp = "portableApp.exe";
					string portableAppName = appDir.Substring(appDir.LastIndexOf('\\') + 1).Replace(" ", ".");
					var versionInfo = FileVersionInfo.GetVersionInfo(appExe);
					string version = string.IsNullOrEmpty(versionInfo.ProductVersion) ? "0" : versionInfo.ProductVersion;
					version = Regex.Replace(version, @"\D", "");
					string fileVersion = !string.IsNullOrEmpty(versionInfo.FileVersion) ? "0" : versionInfo.FileVersion;
					fileVersion = Regex.Replace(fileVersion, @"\D", "");
					string finalVersion = (int.Parse(version) > int.Parse(fileVersion)) ? version : fileVersion;
					portableAppName += ("_" + (finalVersion + "000").Substring(0, 3) + ".exe");
					appDir += "\\";
					Log("version : " + version);
					Log("fileVersion : " + fileVersion);
					Log("finalVersion : " + finalVersion);
					Log("portableAppTemp : " + portableAppTemp);
					Log("portableAppName : " + portableAppName);
					var configSFX = new StreamWriter(portablerFolder + "ConfigSFX.txt");
					configSFX.WriteLine("TempMode");
					configSFX.WriteLine("Silent=2");
					configSFX.WriteLine("Overwrite=2");
					configSFX.WriteLine("Title=Portable App by PortableR");
					configSFX.WriteLine("Setup=" + appName);
					configSFX.Flush();
					configSFX.Close();
					string cmd, parameters;
					try {
						cmd = portablerFolder + "third-party\\" + "rar.exe";
						parameters = "a -ep -ep1 -r -sfx\"" + portablerFolder + "Default.sfx\" -z\"" + portablerFolder + "ConfigSFX.txt\" \"" + appDir + portableAppTemp + "\" \"" + appDir + "*\"";
						Log("SFX Creation", "title");
						Log("starting cmd : " + cmd);
						Log("with params : " + parameters);
						Process sfxProcess = Process.Start(cmd, parameters);
						sfxProcess.WaitForExit();
					} catch (Exception ex) {
						Log("error while producing sfx : " + ex.Message, "error");
					}
					try {
						cmd = portablerFolder + "third-party\\" + "ResourceHacker.exe";
						parameters = "-extract" + " " + "\"" + appExe + "\"" + ", MyProgIcons.rc, icongroup,,";
						Log("Icon Extract", "title");
						Log("starting cmd : " + cmd);
						Log("with params : " + parameters);
						Process extractProcess = Process.Start(cmd, parameters);
						extractProcess.WaitForExit();
					} catch (Exception ex) {
						Log("error while extracting icon : " + ex.Message, "error");
					}
					try {
						Log("Icon Injection", "title");
						var iconFile = new FileInfo(appDir + "Icon_1.ico");
						long iconFileSize = 0;
						if (iconFile.Exists) {
							iconFileSize = new FileInfo(appDir + "Icon_1.ico").Length;
							Log("icon file size  : " + iconFileSize);
						}
						if (iconFileSize > 1) {
							cmd = portablerFolder + "third-party\\" + "ResourceHacker.exe";
							parameters = "-addoverwrite" + " " + "\"" + appDir + portableAppTemp + "\"" + "," + portableAppName + ",Icon_1.ico,ICONGROUP,MAINICON,0";
							Log("starting cmd : " + cmd);
							Log("with params : " + parameters);
							Process injectProcess = Process.Start(cmd, parameters);
							injectProcess.WaitForExit();
						} else {
							Log("no icon injection : icon file not extracted", "error");
							File.Delete(appDir + portableAppName);
							File.Move(appDir + portableAppTemp, appDir + portableAppName);
						}
					} catch (Exception ex) {
						Log("error while injecting icon : " + ex.Message, "error");
					}
					try {
						Log("Temp File Cleaning", "title");
						Log("deleting file : " + portableAppTemp);
						File.Delete(appDir + portableAppTemp);
						Log("deleting file : " + "MyProgIcons.rc");
						File.Delete(appDir + "MyProgIcons.rc");
						string[] icons = Directory.GetFiles(appDir, "*.ico");
						foreach (string icon in icons) {
							Log("deleting file : " + icon);
							File.Delete(icon);
						}
					} catch (Exception ex) {
						Log("error while cleaning temp files : " + ex.Message, "error");
					}
					break;
				case "":
					Log("no action given");
					break;
				default:
					Log("non handled action");
					break;
			}

            // close log
            Log("PortableR end", "end");

            if (hadError)
            {
                System.Windows.Forms.MessageBox.Show("Some error(s) happended, look at " + portablerFolder + logFile);
            }
        }

        static string logPrefix, logSuffix, logSpaces;
        const int logMargin = 18;

        static void Log(string str, string type = "debug")
        {

            logPrefix = "";
            logSuffix = "";

			switch (type) {
				case "error":
					hadError = true;
					logPrefix = "[ ERROR ] ";
					break;
				case "debug":
					logPrefix = "[ debug ] ";
					break;
				case "title":
					logPrefix = "\n\n- ";
					logSuffix = "\n------------------------------";
					break;
				case "end":
					logPrefix = "\n\n=============================\n= ";
					logSuffix = "\n=============================";
					break;
				case "draw":
					logPrefix = " ";
					break;
			}

            int index = str.IndexOf(":");
            if (index > 0)
            {
                logSpaces = new string(' ', logMargin - index);
                str = str.Substring(0, index) + logSpaces + " : " + str.Substring(index + 1);
            }

            logInstance.WriteLine(logPrefix + str + logSuffix);

            logInstance.Flush();

            if (type == "end")
            {
                logInstance.Close();
            }
        }

    }

}
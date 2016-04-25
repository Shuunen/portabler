using System;
using System.Diagnostics;
using System.IO;
using System.Text.RegularExpressions;
using System.Windows.Forms;
using Microsoft.Win32;

namespace PortableR
{
	class Program
	{
		const string logFile = "debug.log";
		static string extractFolder = "%UserProfile%\\Apps\\";
		static StreamWriter logInstance;
		static bool hadError = false;
		static string portablerExe, portablerFolder, portablerBaseFolder, portableAppName;
		const string portableAppTemp = "portableApp.exe";
		static string iconsFolder, sfxFolder, thirdPartyFolder;
		const string extractedIcon = "Icon_1";
		const string fallbackIcon = "store";
		static string appExe, appName, appDir;
		static string cmd, parameters, action;

		static void Main()
		{

			// init
			string[] arguments = Environment.GetCommandLineArgs();
			portablerExe = Application.ExecutablePath; // C:\Projects\portabler\app\bin\PortableR.exe
			portablerFolder = Path.GetDirectoryName(Application.ExecutablePath) + "\\";  // C:\Projects\portabler\app\bin\
			portablerBaseFolder = portablerFolder.Replace("app\\bin\\", ""); // C:\Projects\portabler\
			iconsFolder = portablerBaseFolder + "icons" + "\\";  // C:\Projects\portabler\icons\
			sfxFolder = portablerBaseFolder + "sfx" + "\\";  // C:\Projects\portabler\sfx\
			thirdPartyFolder = portablerBaseFolder + "third-party" + "\\";  // C:\Projects\portabler\third-party\
			extractFolder = Environment.ExpandEnvironmentVariables(extractFolder); // C:\User\YOU\Apps\
			action = arguments.Length > 1 ? arguments[1] : "";
			logInstance = new StreamWriter(portablerBaseFolder + logFile);
			logInstance.AutoFlush = true;			
			//action += "Cancel";

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
			for (int i = 0; i < arguments.Length; i++) {
				Log("argument[" + i + "] : " + arguments[i]);
			}
			Log("portablerExe : " + portablerExe); // C:\Projects\portabler\app\bin\PortableR.exe
			Log("portablerFolder : " + portablerFolder); // C:\Projects\portabler\app\bin\
			Log("portableAppTemp : " + portableAppTemp); // portableApp.exe
			Log("iconsFolder : " + iconsFolder); // C:\Projects\portabler\icons\
			Log("sfxFolder : " + sfxFolder); // C:\Projects\portabler\sfx\
			Log("thirdPartyFolder : " + thirdPartyFolder); // C:\Projects\portabler\third-party\
			Log("extractFolder : " + extractFolder); // C:\User\YOU\Apps\
			Log("action : " + action);
			
			Log(action, "title");
			switch (action) {
					
				case "install":
					var key = Microsoft.Win32.Registry.ClassesRoot;
					try {
						key = key.CreateSubKey("exefile\\shell\\PortableR");
						key.SetValue("MUIVerb", "PortableR");
						key.SetValue("SubCommands", "PortableRcreate;PortableRextract");
						key.SetValue("icon", portablerExe);
					} catch (Exception ex) {
						Log("error during writing to root reg : " + ex, "error");
					} finally {
						key.Close();
					}
					CreateSubCommand("Create package", "create", "windows");
					CreateSubCommand("Extract package", "extract", "sharethis");
					break;
					
				case "create":
					try {
						if (arguments.Length > 2) {
							appExe = arguments[2];							
						} else {
							Log("no targeted app to create app, missing third parameter", "error");
							break;
						}
						appName = Path.GetFileName(appExe);
						appDir = Path.GetDirectoryName(appExe) + "\\";
						Log("appExe : " + appExe);
						Log("appName : " + appName);
						Log("appDir : " + appDir);
						string appDirName = Path.GetFileName(Path.GetDirectoryName(appDir));
						Log("appDirName : " + appDirName);
						if (Regex.IsMatch(appDirName, "^([0-9])")) {
							Log("appDir invalid : appDir start with a number, this generate errors with RessourceHacker icon injection, please rename it", "error");
							return;
						}
						
						var dir = Path.GetDirectoryName(appExe);
						Log("dir : " + dir);
						portableAppName = dir.Substring(dir.LastIndexOf('\\') + 1).Replace(" ", ".");
						// portableAppName = portableAppName.ToLower(); // lowercase
						portableAppName = char.ToUpper(portableAppName[0]) + portableAppName.Substring(1);	// uppercase first letter
						Log("portableAppName : " + portableAppName);
						var versionInfo = FileVersionInfo.GetVersionInfo(appExe);
						var version = string.IsNullOrEmpty(versionInfo.ProductVersion) ? "000" : versionInfo.ProductVersion;
						version = Regex.Replace(version, @"\D", "") + "000";
						version = version.Substring(0, 3);
						Log("version : " + version);
						var fileVersion = string.IsNullOrEmpty(versionInfo.FileVersion) ? "000" : versionInfo.FileVersion;
						fileVersion = Regex.Replace(fileVersion, @"\D", "") + "000";
						fileVersion = fileVersion.Substring(0, 3);
						Log("fileVersion : " + fileVersion);
						var finalVersion = (int.Parse(version) > int.Parse(fileVersion)) ? version : fileVersion;
						finalVersion = (finalVersion + "000").Substring(0, 3);
						// finalVersion = finalVersion == "000" ? "100" : finalVersion;
						portableAppName += ("_" + finalVersion + ".exe");
						Log("finalVersion : " + finalVersion);
					} catch (Exception ex) {
						Log("error while defining variables : " + ex, "error");
					}
					try {
						var configSFX = new StreamWriter(sfxFolder + "conf.txt");
						configSFX.WriteLine("TempMode");
						configSFX.WriteLine("Silent=2");
						configSFX.WriteLine("Overwrite=2");
						configSFX.WriteLine("Title=Portable App by PortableR");
						configSFX.WriteLine("Setup=" + appName);
						configSFX.Flush();
						configSFX.Close();
					} catch (Exception ex) {
						Log("error while creating sfx config : " + ex, "error");
					}
					try {
						cmd = thirdPartyFolder + "rar.exe";
						parameters = "a -ep -ep1 -r -sfx\"" + sfxFolder + "base.sfx\" -z\"" + sfxFolder + "conf.txt\" \"" + appDir + portableAppTemp + "\" \"" + appDir + "*\"";
						Log("SFX Creation", "title");
						Log("starting cmd : " + cmd);
						Log("with params : " + parameters);
						Process sfxProcess = Process.Start(cmd, parameters);
						sfxProcess.WaitForExit();
					} catch (Exception ex) {
						Log("error while producing sfx : " + ex, "error");
					}
					
					Log("Icon Extract", "title");
					string iconFilePath = GetIconFromPaths(new string[] {
						appDir + "icon.ico",
						appDir + extractedIcon + ".ico"
					});
					if (iconFilePath == "") {
						try {
							cmd = thirdPartyFolder + "ResourceHacker.exe";
							parameters = "-extract" + " " + "\"" + appExe + "\"" + ", MyProgIcons.rc, icongroup,,";
							Log("starting cmd : " + cmd);
							Log("with params : " + parameters);
							Process extractProcess = Process.Start(cmd, parameters);
							extractProcess.WaitForExit();
						} catch (Exception ex) {
							Log("error while extracting icon : " + ex, "error");
						}
					} else {
						Log("No extraction needed : icon already in place");
					}
					
					try {
						Log("Icon Injection", "title");
						if (iconFilePath == "") {
							iconFilePath = GetIconFromPaths(new string[] {
								appDir + extractedIcon + ".ico",
								iconsFolder + fallbackIcon + ".ico"
							});
						}
						if (iconFilePath != "") {
							cmd = thirdPartyFolder + "ResourceHacker.exe";
							parameters = "-addoverwrite" + " " + "\"" + appDir + portableAppTemp + "\"" + "," + portableAppName + "," + "\"" + iconFilePath + "\"" + ",ICONGROUP,MAINICON,0";
							Log("starting cmd : " + cmd);
							Log("with params : " + parameters);
							Process injectProcess = Process.Start(cmd, parameters);
							injectProcess.WaitForExit();
							var file = new FileInfo(appDir + portableAppName);
							if (!file.Exists) {
								Log("output app file not found : icon injection failed with target app extracted icon, trying with default icon");
								parameters = parameters.Replace(iconFilePath, iconsFolder + fallbackIcon + ".ico");
								Log("starting cmd : " + cmd);
								Log("with params : " + parameters);
								Process fallbackInjectProcess = Process.Start(cmd, parameters);
								fallbackInjectProcess.WaitForExit();
							}
						} else {
							Log("no icon injection : icon file not found", "error");
							File.Delete(appDir + portableAppName);
							File.Move(appDir + portableAppTemp, appDir + portableAppName);
						}
					} catch (Exception ex) {
						Log("error while injecting icon : " + ex, "error");
					}
					try {
						var file = new FileInfo(appDir + portableAppName);
						if (!file.Exists) {
							Log("portable app has not been generated, cleaning step aborted", "error");
						} else {
							Log("Temp File Cleaning", "title");
							DeleteIfExists(appDir + portableAppTemp);
							DeleteIfExists(appDir + "MyProgIcons.rc");
							string[] icons = Directory.GetFiles(appDir, "*.ico");
							foreach (string icon in icons) {
								DeleteIfExists(icon);
							}
						}
					} catch (Exception ex) {
						Log("error while cleaning temp files : " + ex, "error");
					}
					break;
					
				case "extract":
					try {
						string pathsStr;
						if (arguments.Length > 2) {
							pathsStr = arguments[2];							
						} else {
							Log("no targeted app(s) to extract app, missing third parameter", "error");
							break;
						}
						pathsStr += '|';
						string[] paths = pathsStr.Split('|');
						foreach (string path in paths) {
							if (path.Length < 1) {
								Log("Skip empty path...");
								continue;
							}
							appExe = path;
							appName = Path.GetFileName(appExe);
							appDir = Path.GetDirectoryName(appExe) + "\\";
							Log("App Extract", "title");
							Log("appExe : " + path);
							Log("appName : " + appName);
							Log("appDir : " + appDir);
							string folderName = appName.Substring(0, appName.LastIndexOf("_", StringComparison.Ordinal)) + "\\";
							cmd = thirdPartyFolder + "rar.exe";
							parameters = "x  " + "\"" + appExe + "\"" + " " + "\"" + extractFolder + folderName + "\"";							
							Log("starting cmd : " + cmd);
							Log("with params : " + parameters);
							Process extractProcess = Process.Start(cmd, parameters);
							extractProcess.WaitForExit();	
						}
					} catch (Exception ex) {
						Log("error while extracting app : " + ex, "error");
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

			if (hadError) {
				MessageBox.Show("Some error(s) happened, look at " + portablerFolder + logFile);
			}
		}

		static string GetIconFromPaths(string[] paths)
		{
			string selectedPath = "";
			foreach (var path in paths) {
				Log("icon file test : " + path);
				var file = new FileInfo(path);
				Log("icon file exists : " + file.Exists.ToString());
				if (file.Exists && file.Length != 0) {
					Log("icon file size : " + file.Length);
					Log("icon file selected");
					selectedPath = path;
					break;
				}
			}
			return selectedPath;
		}
		
		static void CreateSubCommand(string title, string command, string icon = "github")
		{
			Log("CreateSubCommand", "title");
			Log("title : " + title);
			Log("command : " + command);
			Log("icon : " + icon);
			string iconPath = iconsFolder + icon + ".ico";
			Log("iconPath : " + iconPath);
			
			var key = Microsoft.Win32.Registry.LocalMachine;
			// var baseReg = RegistryKey.OpenBaseKey(RegistryHive.LocalMachine,RegistryView.Registry64);
			try {
				key = key.CreateSubKey("SOFTWARE\\Microsoft\\Windows\\CurrentVersion\\Explorer\\CommandStore\\shell\\PortableR" + command);
				key.SetValue("", title);
				key.SetValue("icon", iconPath);
				key = key.CreateSubKey("command");
				key.SetValue("", thirdPartyFolder + "Elevate.exe" + " " + portablerExe + " " + command + " \"%0\"");
			} catch (Exception ex) {
				Log("error during writing \"" + command + "\" command to local machine reg : " + ex, "error");
			} finally {
				key.Close();
			}
		}
		
		static void DeleteIfExists(string path)
		{
			string filename = Path.GetFileName(path);
			var file = new FileInfo(path);
			if (file.Exists) {
				Log("file exists : deleting file (" + path + ")");
				File.Delete(path);
			} else {
				Log("file do not exists : skip delete (" + path + ")");
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
					logPrefix = "\n\n[ ERROR ] ";
					logSuffix = "\n\n";
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
			
			int index = str.IndexOf(":", StringComparison.Ordinal);
			if (index > 0 && (logMargin - index) > 0) {
				try {
					logSpaces = new string(' ', logMargin - index);
					str = str.Substring(0, index) + logSpaces + " : " + str.Substring(index + 1);
				} catch (Exception ex) {
					str = str + "\n\n" + ex + "\n\n";
				}
			}
			
			str = logPrefix + str + logSuffix;
			
			logInstance.WriteLine(str);

			logInstance.Flush();

			if (type == "end") {
				logInstance.Close();
			}
		}
		
		static string RemoveSpecialCharacters(string str)
		{
			return Regex.Replace(str, "[^a-zA-Z0-9_.]+", "", RegexOptions.Compiled);
		}
	}
}
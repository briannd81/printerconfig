////////////////////////////////////////////////////////////////////////////
// Chinh Dang
// Purpose: Self-manage printing application for laptop users
// Version: 3
//
////////////////////////////////////////////////////////////////////////////
using System;
using System.Diagnostics;
using System.Globalization;
using System.Collections.Generic;
using System.Text;
using System.Management;
using System.IO;
using System.Threading;
using System.Reflection;

[assembly: AssemblyVersion("3.0.0.0")]

namespace ConsoleApplication1
{
    /// <summary>
    /// Main Class
    /// </summary>
    class Program
    {
        static void Main(string[] args)
        {
            //print a divider line to log to indicate program execution start
            Logger.logging("-------------------------------------------------------");

            //some brackground info about the printer
            //--------------------------------------
            //get printer info: prncnfg.vbs -g -p printer
            //  printer -> port     -> driver
            //  hplaser -> pBBBBBBa -> HP LaserJet 6P
            //  MX611   -> pBBBBBBm -> Lexmark Universal v2 PS3
            //get port info: prnport.vbs -g -r port
            //  port     -> protocol -> host address -> [queue|port number]
            //  pBBBBBBa -> LPR ->   -> p777777a     -> p777777a
            //  pBBBBBBm -> RAW ->   -> p777777m     -> 9100
            //--------------------------------------

            //setup some local variables to use in main
            string opt = null;
            string dummyStr = null;
            string defaultPrinter = null;
            string defaultPort = null; //of default printer
            string defaultPortHost = null; //to the user this is the their default printer
            string branchNumber = null; //defaultPortHost=p123456a -> branchNumber=123456
            int optMin = 0;
            int optMax = 0;
            int dummyInt = 0;

            //create an instance of class Program to access the member function/variable
            Program p = new Program();

            try
            {
                while (true)
                {
                    //reset
                    Logger.logging("main(): Reset default printer, port, host, and branch number to null");
                    defaultPrinter = null;
                    defaultPort = null;
                    defaultPortHost = null;
                    branchNumber = null;

                    //get default printer port
                    Logger.logging("main(): Get default printer");
                    defaultPrinter = Printer.getDefaultPrinter();

                    if (!string.IsNullOrEmpty(defaultPrinter) && isDefaultPrinterValid(defaultPrinter))
                    {
                        Logger.logging("main(): Default printer " + defaultPrinter + " is valid");

                        //get default port
                        Logger.logging("main(): Get port of default printer");
                        defaultPort = Printer.getDefaultPort();

                        if (!string.IsNullOrEmpty(defaultPort) && isPortValidForPrinter(defaultPrinter, defaultPort))
                        {
                            Logger.logging("main(): port " + defaultPort + " is valid");

                            //get default printer port host
                            Logger.logging("main(): Get port host of default printer");
                            defaultPortHost = Printer.getPortHost(defaultPort);

                            //get the branch number if port host is valid
                            //  valid if 8 char, starts with p, ends with [a,b,g,h,m], and index[1,6] is a number
                            Logger.logging("main(): Check if port host is valid, i.e. p123456a");
                            if (!string.IsNullOrEmpty(defaultPortHost) && isHostValidForPort(defaultPort, defaultPortHost))
                            {
                                Logger.logging("main(): Port host " + defaultPortHost + " is valid");

                                //extract branch number from port host
                                Logger.logging("main(): Extract branch number from port host");
                                if (int.TryParse(defaultPortHost.Substring(1, 6), out dummyInt))
                                {
                                    branchNumber = defaultPortHost.Substring(1, 6);
                                    Logger.logging("main(): Branch number is set to " + branchNumber);
                                }
                                else
                                {
                                    Logger.logging("main(): Unable to extract branch number from port host");
                                }
                            }
                            else
                            {
                                //assign null if port host invalid
                                Logger.logging("main(): Port host " + defaultPortHost + " is invalid");
                            }
                        }
                        else
                        {
                            Logger.logging("main(): Port " + defaultPort + " is invalid");
                        }
                    }
                    else
                    {
                        Logger.logging("main(): Default printer " + defaultPrinter + " is invalid");
                    }

                    //log default printer, port, and host
                    Logger.logging("main(): Default printer: " + defaultPrinter);
                    Logger.logging("main(): Default port: " + defaultPort);
                    Logger.logging("main(): Default host: " + defaultPortHost);
                    Logger.logging("main(): Branch number: " + branchNumber);

                    //display main menu
                    //////////////////////////////////////
                    //Main Options:
                    //1. Change Branch Number
                    //2. Change Default Printer
                    //////////////////////////////////////
                    Console.Clear();

                    //display port host last char if port host is valid
                    if (!string.IsNullOrEmpty(defaultPortHost) && isHostValidForPort(defaultPort, defaultPortHost))
                    {
                        displayMainMenu(defaultPortHost[7].ToString(), branchNumber);
                    }
                    else
                    {
                        displayMainMenu(defaultPrinter, branchNumber);
                    }

                    //set optMin and optMax
                    optMin = 1;
                    optMax = 3;

                    //prompt user to select an option from main
                    printNewLine();
                    opt = getSelectedOption();

                    //make sure input is an int
                    if (int.TryParse(opt, out dummyInt))
                    {
                        if (validateOption(opt, optMin, optMax) == true)
                        {
                            printNewLine();
                            dummyInt = Convert.ToInt16(opt);

                            //Main Menu -> 1. Change Branch Number
                            if (dummyInt == 1)
                            {
                                Logger.logging("main(): User selected option \"1. Change Branch Number\"");
                                Console.Write(" Enter new branch number: ");
                                dummyStr = Console.ReadLine();
                                if (isBranchValid(dummyStr))
                                {
                                    Logger.logging("main(): Branch number " + dummyStr + " entered is valid");

                                    //create dictionary of port and host using the branch number
                                    Dictionary<string, string> portHost = new Dictionary<string, string>();
                                    portHost.Add("pBBBBBBa", "p" + dummyStr + "a");
                                    portHost.Add("pBBBBBBb", "p" + dummyStr + "b");
                                    portHost.Add("pBBBBBBg", "p" + dummyStr + "g");
                                    portHost.Add("pBBBBBBh", "p" + dummyStr + "h");
                                    portHost.Add("pBBBBBBm", "p" + dummyStr + "m");

                                    //set port host status flag
                                    bool setHostFail = false;

                                    //update host for ports a,b,g,h,m
                                    //i.e. p111111a -> p222222a, p111111m -> p222222m, etc...
                                    Logger.logging("main(): Update all hosts with branch number " + dummyStr);
                                    foreach (KeyValuePair<string, string> port in portHost)
                                    {
                                        if (Printer.setPortHost(port.Key, port.Value))
                                        {
                                            Logger.logging("main(): Successfully update host " + port.Value + " for port " + port.Key);
                                        }
                                        else
                                        {
                                            Logger.logging("main(): Unable to update host " + port.Value + " for port " + port.Key);
                                            setHostFail = true;
                                        }
                                    }

                                    //print set port host status
                                    if (setHostFail)
                                    {
                                        printNewLine();
                                        printToConsole(" - Failed to set branch number");
                                    }
                                    else
                                    {
                                        printNewLine();
                                        printToConsole(" - Successfully changed branch number");
                                    }

                                }
                                else
                                {
                                    printNewLine();
                                    printToConsole(" - Invalid branch number");
                                    Logger.logging("main(): Branch number " + dummyStr + " entered is invalid");
                                }

                                printNewLine();
                                keyToContinue();
                                continue;
                            }
                            //Main Menu -> 2. Change Default Printer
                            else if (dummyInt == 2)
                            {
                                Logger.logging("main(): User selected option \"2. Change Default Printer\"");

                                //get list of configured printers
                                List<string> printerList = new List<string>(Printer.getPrinterList());

                                //create dictionary that user can select as default
                                Dictionary<int, string> defaultPrinterDict = new Dictionary<int, string>();
                                defaultPrinterDict.Add(1, "hplaser");
                                defaultPrinterDict.Add(2, "micr");
                                defaultPrinterDict.Add(3, "MX611");
                                defaultPrinterDict.Add(4, "Other");

                                //create another dictionary but exclude known printers (hplaser, micr, etc...)
                                Dictionary<int, string> otherPrinterDict = new Dictionary<int, string>();
                                for (int i=0, count=1; i < printerList.Count; i++)
                                {
                                    if (!string.IsNullOrEmpty(printerList[i]) &&
                                        (!printerList[i].Equals("hplaser", StringComparison.OrdinalIgnoreCase) &&
                                        !printerList[i].Equals("repprt", StringComparison.OrdinalIgnoreCase) &&
                                        !printerList[i].Equals("locprt", StringComparison.OrdinalIgnoreCase) &&
                                        !printerList[i].Equals("micr", StringComparison.OrdinalIgnoreCase) &&
                                        !printerList[i].Equals("MX611", StringComparison.OrdinalIgnoreCase)))
                                    {
                                        otherPrinterDict.Add(count, printerList[i]);
                                        count++;
                                    }
                                }

                                //display printers to make default
                                //1. hplaser
                                //2. micr
                                //3. MX611
                                //4. Other
                                displayDefaultPrinterSelection(defaultPrinterDict);

                                printNewLine();
                                printToConsole(" Note: Option 1-3 are branch printers");
                                printNewLine();
                                Console.Write(" Enter line number [1-" + defaultPrinterDict.Count + "]: ");
                                dummyStr = Console.ReadLine();
                                //User selected "1. hplaser", "2. micr", or "3. MX611"
                                if (int.TryParse(dummyStr, out dummyInt) && (dummyInt >= 1 && dummyInt < defaultPrinterDict.Count))
                                {                                    
                                    //make sure the key-value exist
                                    if (defaultPrinterDict.ContainsKey(dummyInt))
                                    {
                                        Logger.logging("main(): User selected printer " + defaultPrinterDict[dummyInt] + " as default");
                                        //make sure printer selected is not default before proceed
                                        if (!string.IsNullOrEmpty(defaultPrinter) &&
                                            defaultPrinter.Equals(defaultPrinterDict[dummyInt], StringComparison.OrdinalIgnoreCase))
                                        {
                                            printNewLine();
                                            printToConsole(" - Selected printer is already the default");
                                        }
                                        else if (Printer.setDefaultPrinter(defaultPrinterDict[dummyInt]))
                                        {
                                            printNewLine();
                                            //printToConsole(" - Printer " + defaultPrinterDict[dummyInt] + " has been set to default");
                                            printToConsole(" - Successfully changed default printer");
                                            Logger.logging("main(): Change default printer successful");
                                        }
                                        else
                                        {
                                            printNewLine();
                                            //printToConsole(" - Unable to change default printer to " + defaultPrinterDict[dummyInt]);
                                            printToConsole(" - Failed to change default printer");
                                            Logger.logging("main(): Failed to change default printer");
                                        }
                                    }
                                    else
                                    {
                                        Logger.logging("main(): Unable to locate value for defaultPrinterDict[" + dummyInt + "] key");
                                        printToConsole(" - Unable to determine new default printer name");
                                    }
                                }
                                else if (int.TryParse(dummyStr, out dummyInt) && (dummyInt == defaultPrinterDict.Count))
                                {
                                    //User selected "4. Other"
                                    Logger.logging("User selected to change Other default printer");

                                    //display other printers to make default
                                    printNewLine();
                                    displayDefaultPrinterSelection(otherPrinterDict);

                                    //If there are other printers
                                    if (otherPrinterDict.Count > 0)
                                    {
                                        printNewLine();
                                        if (otherPrinterDict.Count == 1)
                                        {
                                            Console.Write(" Enter line number: ");
                                        }
                                        else
                                        {
                                            Console.Write(" Enter line number [1-" + otherPrinterDict.Count + "]: ");
                                        }
                                        dummyStr = Console.ReadLine();

                                        //Make sure option selected is valid
                                        if (int.TryParse(dummyStr, out dummyInt) &&
                                            (dummyInt >= 1 && dummyInt <= otherPrinterDict.Count))
                                        {
                                            //make sure the key-value exist
                                            if (otherPrinterDict.ContainsKey(dummyInt))
                                            {
                                                Logger.logging("main(): User selected printer " + otherPrinterDict[dummyInt] + " as default");
                                                //make sure printer selected is not default before proceed
                                                if (!string.IsNullOrEmpty(defaultPrinter) &&
                                                    defaultPrinter.Equals(otherPrinterDict[dummyInt], StringComparison.OrdinalIgnoreCase))
                                                {
                                                    printNewLine();
                                                    printToConsole(" - Selected printer is already the default");
                                                }
                                                else if (Printer.setDefaultPrinter(otherPrinterDict[dummyInt]))
                                                {
                                                    printNewLine();
                                                    //printToConsole(" - Printer " + otherPrinterDict[dummyInt] + " has been set to default");
                                                    printToConsole(" - Successfully changed default printer");
                                                    Logger.logging("main(): Change default printer successful");
                                                }
                                                else
                                                {
                                                    printNewLine();
                                                    //printToConsole(" - Unable to change default printer to " + otherPrinterDict[dummyInt]);
                                                    printToConsole(" - Failed to change default printer");
                                                    Logger.logging("main(): Failed to change default printer");
                                                }
                                            }
                                            else
                                            {
                                                Logger.logging("main(): Unable to locate value for otherPrinterDict[" + dummyInt + "] key");
                                                printToConsole(" - Unable to determine new default printer name");
                                            }
                                        }
                                        else
                                        {
                                            printNewLine();
                                            printToConsole(" - Invalid input");
                                            Logger.logging("main(): User entered invalid other default printer option " + dummyStr);
                                        }
                                    }
                                    else
                                    {
                                        //printNewLine();
                                        printToConsole(" - There are no other printers");
                                    }
                                }
                                else
                                {
                                    printNewLine();
                                    printToConsole(" - Invalid input");
                                    Logger.logging("main(): User entered invalid default printer option " + dummyStr);
                                }
                                printNewLine();
                                keyToContinue();
                            }
                            else if (dummyInt == 3)
                            {
                                Logger.logging("main(): User exit application");
                                break;
                            }
                            else
                            {
                                printNewLine();
                                printInvalidOptionMsg();
                                printNewLine();
                                keyToContinue();
                                continue;
                            }
                        }
                        else
                        {
                            printNewLine();
                            printInvalidOptionMsg();
                            printNewLine();
                            keyToContinue();
                            continue;
                        }
                    }
                    else
                    {
                        //input not an int
                        if (opt.Equals("show admin", StringComparison.OrdinalIgnoreCase))
                        {
                            displayAdminMenu();
                            continue;
                        }
                        else if (opt.Equals("reset config", StringComparison.OrdinalIgnoreCase))
                        {
                            Logger.logging("main(): User reset printer configuration to default");
                            Printer.resetConfig();
                        }
                        else if (opt.Equals("cleanup port", StringComparison.OrdinalIgnoreCase))
                        {
                            Logger.logging("main(): User cleanup non-standard port");
                            Printer.cleanupPort();
                        }
                        else
                        {
                            printNewLine();
                            printInvalidOptionMsg();
                            printNewLine();
                            keyToContinue();
                            continue;
                        }
                    }

                    continue;
                }
            }
            catch (Exception e)
            {
                Logger.logging("main() Exception: " + e.Message);
            }
        }
        /// <summary>
        /// Return true if default printer is hplaser, locprt, repprt, micr, or MX611
        /// </summary>
        /// <returns></returns>
        public static bool isDefaultPrinterValid(string printer)
        {
            try
            {
                if (string.IsNullOrEmpty(printer))
                {
                    return false;
                }
                else if ((string.Equals("hplaser", printer, StringComparison.OrdinalIgnoreCase)) ||
                    (string.Equals("repprt", printer, StringComparison.OrdinalIgnoreCase)) ||
                    (string.Equals("locprt", printer, StringComparison.OrdinalIgnoreCase)) ||
                    (string.Equals("micr", printer, StringComparison.OrdinalIgnoreCase)) ||
                    (string.Equals("MX611", printer, StringComparison.OrdinalIgnoreCase)))
                {
                    return true;
                }
                else
                {
                    return false;
                }
            }
            catch (Exception e)
            {
                Logger.logging("isDefaultPrinterValid() Exception: " + e.Message);
            }
            return false;
        }
        /// <summary>
        /// Return true if port name is valid - pBBBBBB[a|b|m]
        /// </summary>
        /// <param name="port"></param>
        /// <returns></returns>
        public static bool isPortNameValid(string port)
        {
            try
            {
                //create list of known ports
                List<string> portList = new List<string>();
                portList.Add("pBBBBBBa");
                portList.Add("pBBBBBBb");
                portList.Add("pBBBBBBm");

                if (!string.IsNullOrEmpty(port) && portList.Contains(port))
                {
                    return true;
                }
                else
                {
                    return false;
                }
            }
            catch (Exception e)
            {
                Logger.logging("isPortNameValid() Exception: " + e.Message);
            }
            return false;
        }
        /// <summary>
        /// Return true if the port is valid for the printer
        /// </summary>
        /// <returns></returns>
        public static bool isPortValidForPrinter(string printer, string port)
        {
            try
            {
                Dictionary<string, string> validPrinterPort = new Dictionary<string, string>();
                validPrinterPort.Add("hplaser", "pBBBBBBa");
                validPrinterPort.Add("repprt", "pBBBBBBa");
                validPrinterPort.Add("locprt", "pBBBBBBa");
                validPrinterPort.Add("micr", "pBBBBBBb");
                validPrinterPort.Add("MX611", "pBBBBBBm");

                //return true if both match printer and port
                if (!string.IsNullOrEmpty(printer) && !string.IsNullOrEmpty(port) &&
                    validPrinterPort.ContainsKey(printer) && validPrinterPort[printer].Equals(port))
                {
                    return true;
                }
                else
                {
                    return false;
                }
            }
            catch (Exception e)
            {
                Logger.logging("isPortValidForPrinter() Exception: " + e.Message);
            }
            return false;
        }
        /// <summary>
        /// Return true if port host is valid for port
        /// </summary>
        /// <returns></returns>
        public static bool isHostValidForPort(string port, string host)
        {
            try
            {
                int dummyInt = 0;
                // valid host: p111111a
                //  host[0] = p
                //  host[7] = a,b,m
                //  host[1,6] = all digit
                // port[7] = host[7] (i.e. port pBBBBBBa and host p111111a)
                if (!string.IsNullOrEmpty(host) && (host.Length == 8) && 
                    (string.Equals("p", host[0].ToString(), StringComparison.OrdinalIgnoreCase)) &&
                    ("abm".Contains(host[7].ToString())) && int.TryParse(host.Substring(1, 6), out dummyInt) &&
                    !string.IsNullOrEmpty(port) && isPortNameValid(port) &&
                    string.Equals(port[7].ToString(), host[7].ToString()))
                {
                    return true;
                }
                else
                {
                    return false;
                }
            }
            catch (Exception e)
            {
                Logger.logging("isHostValidForPort() Exception: " + e.Message);
            }

            return false;
        }
        /// <summary>
        /// Show printer options
        /// </summary>
        public static void displayDefaultPrinterSelection(Dictionary<int, string> printer)
        {
            try
            {
                Dictionary<string, string> printerNickName = new Dictionary<string, string>();
                printerNickName.Add("hplaser", "Printer a");
                printerNickName.Add("repprt", "Printer a");
                printerNickName.Add("locprt", "Printer a");
                printerNickName.Add("micr", "Printer b (Check Printer)");
                printerNickName.Add("MX611", "Printer m (Multi-Functional Printer)");

                foreach (KeyValuePair<int, string> p in printer)
                {
                    if (printerNickName.ContainsKey(p.Value))
                    {
                        printToConsole(" " + p.Key + ". " + printerNickName[p.Value]);
                    }
                    else
                    {
                        printToConsole(" " + p.Key + ". " + p.Value);
                    }
                }
            }
            catch (Exception e)
            {
                Logger.logging("displayDefaultPrinterSelection() Exception: " + e.Message);
            }

        }
        /// <summary>
        /// Show the admin menu screen
        /// </summary>
        public static void displayAdminMenu()
        {
            try
            {
                Console.Clear();
                printNewLine();
                printToConsole(" Administrative Menu");
                printNewLine();
                printToConsole(" - reset config: reset all printers and ports to default");
                printNewLine();
                printToConsole(" - cleanup port: remove all non-standard ports");
                printNewLine();
                keyToContinue();
            }
            catch (Exception e)
            {
                Logger.logging("displayAdminMenu() Exception: " + e.Message);
            }
        }
        /// <summary>
        /// Return true if branch is valid (6 digits)
        /// </summary>
        /// <param name="str"></param>
        /// <returns></returns>
        public static bool isBranchValid(string str)
        {
            try
            {
                int dummy = -1;
                if (!string.IsNullOrEmpty(str) && int.TryParse(str, out dummy))
                {
                    if (str.Length == 6)
                    {
                        return true;
                    }
                    else
                    {
                        return false;
                    }
                }
            }
            catch (Exception e)
            {
                Logger.logging("isBranchValid() Exception: " + e.Message);
            }

            //return false here
            return false;
        }

        /// <summary>
        /// Print "enter c to cancel" message to console
        /// </summary>
        public static void printCancelInputMsg()
        {
            printToConsole("*** Note: enter c to cancel");
        }
        /// <summary>
        /// Print "invalid option selected" message to console
        /// </summary>
        public static void printInvalidOptionMsg()
        {
            printToConsole(" - Invalid option selected");
        }
        /// <summary>
        /// Return true if it's a laptop
        /// </summary>
        /// <returns></returns>
        public static bool isALaptop()
        {
            int chassistype = -1;
            try
            {

                //define wmi connection object
                ConnectionOptions co = new ConnectionOptions();

                ManagementPath mPath = new ManagementPath(@"\\.\root\cimv2");

                ManagementScope mScope = new ManagementScope(mPath, co);

                mScope.Connect();

                ObjectQuery oQuery = new ObjectQuery("SELECT * FROM Win32_SystemEnclosure");

                ManagementObjectSearcher mObjSearcher = new ManagementObjectSearcher(mScope, oQuery);

                ManagementObjectCollection getObjColl = mObjSearcher.Get();

                foreach (ManagementObject mObj in getObjColl)
                {
                    foreach (int i in (UInt16[])(mObj["ChassisTypes"]))
                    {
                        chassistype = i;
                    }
                }


                //chassistype(3 -> desktop, 8 -> portable (Dell laptop return this), 9 -> laptop, 10 -> notebook)
                if (chassistype == 3)
                {
                    return false;
                }
                else if ((chassistype == 8) || (chassistype == 9) || (chassistype == 10))
                {
                    return true;
                }
                else
                {
                    return false;
                }
            }
            catch (Exception e)
            {
                Logger.logging("isALaptop() Exception: " + e.Message);
            }
            return false;
        }
        /// <summary>
        /// Check if string parameter contain any whitespace character
        /// </summary>
        /// <param name="value"></param>
        /// <returns>bool</returns>
        public static bool isNullOrWhiteSpace(string value)
        {
            try
            {
                if (!string.IsNullOrEmpty(value))
                {
                    for (int i = 0; i < value.Length; i++)
                    {
                        if (!char.IsWhiteSpace(value[i]))
                        {
                            return false;
                        }
                    }
                }
                return true;
            }
            catch (Exception e)
            {
                Logger.logging("isNullOrWhiteSpace() Exception: " + e.Message);
            }
            return false;
        }
        /// <summary>
        /// Display password entered in masked char "*"
        /// </summary>
        /// <returns>string</returns>
        public static string enterPassword()
        {
            try
            {
                string password = null;
                ConsoleKeyInfo key;

                do
                {
                    key = Console.ReadKey(true);

                    if (key.Key != ConsoleKey.Backspace && key.Key != ConsoleKey.Enter)
                    {
                        password += key.KeyChar;
                        Console.Write("*");
                    }
                    else
                    {
                        if (key.Key == ConsoleKey.Backspace && password.Length > 0)
                        {
                            password = password.Substring(0, (password.Length - 1));
                            Console.Write("\b \b");
                        }
                    }
                } while (key.Key != ConsoleKey.Enter);

                //return password
                return password;
            }
            catch (Exception e)
            {
                Logger.logging("enterPassword() Exception: " + e.Message);
            }
            return null;
        }
        /// <summary>
        /// Display main menu
        /// </summary>
        /// <param name="domain"></param>
        /// <param name="username"></param>
        /// <param name="password"></param>
        /// <param name="workstation"></param>
        public static void displayMainMenu(string defaultPortHost, string branchNumber)
        {
            try
            {
                //menu string and length - use to determine appropriate alignment
                string hDiv = "-----------------------------------------------------";
                int hDivAlign = hDiv.Length + 1;
                string menuTitle = "OneMain Laptop Printing Assistant";
                int mtAlign = ((hDivAlign + menuTitle.Length) / 2) - 1;
                string wksTitle = " 1. Change Branch Number [" + branchNumber + "]";
                int wtAlign = wksTitle.Length;
                string processTitle = " 2. Change Default Printer ";
                int dphMaxLength = hDiv.Length - processTitle.Length - 4;
                if (string.IsNullOrEmpty(defaultPortHost) || 
                    (!string.IsNullOrEmpty(defaultPortHost) && (defaultPortHost.Length < dphMaxLength)))
                {
                    processTitle = processTitle + "[" + defaultPortHost + "]";
                }
                else
                {
                    processTitle = processTitle + "[" + defaultPortHost.Substring(0, dphMaxLength - 3) + "..]";
                }
                int protAlign = processTitle.Length;
                string exitTitle = " 3. Exit";
                int exitAlign = exitTitle.Length;

                printNewLine();
                Console.WriteLine(string.Format("{0," + hDivAlign + "}", hDiv));
                Console.WriteLine(string.Format("{0,2}{1," + mtAlign + "}{2," + (hDivAlign - mtAlign - 2) + "}", "|", menuTitle, "|"));
                Console.WriteLine(string.Format("{0," + hDivAlign + "}", hDiv));
                Console.WriteLine(string.Format("{0,2}{1," + (hDivAlign - 2) + "}", "|", "|"));
                Console.WriteLine(string.Format("{0,2}{1," + wtAlign + "}{2," + (hDivAlign - wtAlign - 2) + "}", "|", wksTitle, "|"));
                Console.WriteLine(string.Format("{0,2}{1," + (hDivAlign - 2) + "}", "|", "|"));
                Console.WriteLine(string.Format("{0,2}{1," + protAlign + "}{2," + (hDivAlign - protAlign - 2) + "}", "|", processTitle, "|"));
                Console.WriteLine(string.Format("{0,2}{1," + (hDivAlign - 2) + "}", "|", "|"));
                Console.WriteLine(string.Format("{0,2}{1," + exitAlign + "}{2," + (hDivAlign - exitAlign - 2) + "}", "|", exitTitle, "|"));
                Console.WriteLine(string.Format("{0,2}{1," + (hDivAlign - 2) + "}", "|", "|"));
                Console.WriteLine(string.Format("{0," + hDivAlign + "}", hDiv));

                //display login
                //displayLogin(domain, username, password, workstation);
            }
            catch (Exception e)
            {
                Logger.logging("displayMainMenu() Exception: " + e.Message);
            }
        }
        /// <summary>
        /// Press any key to continue
        /// </summary>
        public static void keyToContinue()
        {
            Console.Write(" Press any key to continue... ");
            Console.ReadLine();
            Console.Clear();
        }
        /// <summary>
        /// Print string parameter to console
        /// </summary>
        /// <param name="msg"></param>
        public static void printToConsole(string msg)
        {
            Console.WriteLine(" " + msg);
        }
        /// <summary>
        /// Get user input
        /// </summary>
        /// <returns></returns>
        public static string getSelectedOption()
        {
            Console.Write(" Select an option: ");
            return Console.ReadLine();
        }
        /// <summary>
        /// Print new line
        /// </summary>
        public static void printNewLine()
        {
            Console.WriteLine();
        }
        /// <summary>
        /// Get domain
        /// </summary>
        /// <returns>string</returns>
        static public string getDomain()
        {
            Console.WriteLine("Enter domain: ");
            return Console.ReadLine();
        }
        /// <summary>
        /// Get username
        /// </summary>
        /// <returns></returns>
        static public string getUsername()
        {
            Console.WriteLine("Enter user name: ");
            return Console.ReadLine();
        }
        /// <summary>
        /// Get password
        /// </summary>
        /// <returns></returns>
        static public string getPassword()
        {
            Console.WriteLine("Enter password: ");
            return Console.ReadLine();
        }
        /// <summary>
        /// Validate if options are valid
        /// </summary>
        /// <param name="opt"></param>
        /// <param name="optMin"></param>
        /// <param name="optMax"></param>
        /// <returns>bool</returns>
        public static bool validateOption(string opt, int optMin, int optMax)
        {
            try
            {
                int dummyVal = 0;

                if (int.TryParse(opt, out dummyVal))
                {
                    dummyVal = Convert.ToInt16(opt);
                    if ((dummyVal >= optMin) && (dummyVal <= optMax))
                    {
                        return true;
                    }
                    else
                    {
                        return false;
                    }
                }
                else
                {
                    return false;
                }
            }
            catch (Exception e)
            {
                Logger.logging("validateOption() Exception: " + e.Message);
            }

            return false;
        }
    }

    /// <summary>
    /// Printer Class
    /// </summary>
    public class Printer : IEquatable<Printer>
    {
        public string name;
        public string port;

        public Printer(string name, string port)
        {
            this.name = name;
            this.port = port;
        }

        /// <summary>
        /// Returns true if port is equal
        /// </summary>
        /// <param name="p"></param>
        /// <returns></returns>
        public bool Equals(Printer p)
        {
            if (this.port.Equals(p.port))
            {
                return true;
            }
            else
            {
                return false;
            }
        }
        /// <summary>
        /// Print message to console
        /// </summary>
        /// <param name="msg"></param>
        public static void printToConsole(string msg)
        {
            Console.WriteLine(" " + msg);
        }
        /// <summary>
        /// Print new line
        /// </summary>
        public static void printNewLine()
        {
            Console.WriteLine();
        }
        /// <summary>
        /// Return path %systemroot%\system32\printing_admin_scripts\en-US
        /// </summary>
        /// <returns></returns>
        public static string getPrinterScriptPath()
        {
            try
            {
                string sysRoot = Environment.GetEnvironmentVariable("SystemRoot");
                return sysRoot + "\\system32\\Printing_Admin_Scripts\\en-US\\";
            }
            catch (Exception e)
            {
                Logger.logging("getPrinterScriptPath() Exception: " + e.Message);
            }
            return null;
        }
        /// <summary>
        /// Return %systemroot%\system32
        /// </summary>
        /// <returns></returns>
        public static string getCScriptPath()
        {
            try
            {
                string sysRoot = Environment.GetEnvironmentVariable("SystemRoot");
                return sysRoot + "\\system32\\";
            }
            catch (Exception e)
            {
                Logger.logging("getCScriptPath() Exception: " + e.Message);
            }
            return null;
        }
        /// <summary>
        /// Remove non-default ports that starts with a p, follow by a series of letters, and end with either [a,b,g,h]
        /// </summary>
        public static void cleanupPort()
        {
            bool fail = false;

            try
            {
                //get list of non-standard ports
                List<string> port = Printer.getNonStandardPort();
                if (port.Count > 0)
                {
                    foreach (string prt in port)
                    {
                        Logger.logging("cleanupPort(): Found non-standard port " + prt);
                        Logger.logging("cleanupPort(): Remove non-standard port " + prt);
                        if (removePort(prt))
                        {
                            Logger.logging("cleanupPort(): Successfully remove port " + prt);
                        }
                        else
                        {
                            Logger.logging("cleanupPort(): Failed to remove port " + prt);
                            fail = true;
                        }
                    }

                    //print status message
                    if (fail)
                    {
                        printNewLine();
                        printToConsole(" - Failed to remove non-standard ports");
                    }
                    else
                    {
                        printNewLine();
                        printToConsole(" - Successfully remove non-standard ports");
                    }

                }
                else
                {
                    Logger.logging("cleanupPort(): Non-standard port not found");
                    printNewLine();
                    printToConsole(" - There are no non-standard ports to remove");
                }

                printNewLine();
                Program.keyToContinue();

            }
            catch (Exception e)
            {
                Logger.logging("cleanupPort() Exception: " + e.Message);
            }
        }
        /// <summary>
        /// Reset printer and port configuration to default
        /// </summary>
        public static void resetConfig()
        {

            try
            {
                //print message that it may take a while
                printNewLine();
                printToConsole(" - Reset printer configuration to default. This may take a while");

                //create list of printers
                List<string> printer = new List<string>();
                printer.Add("hplaser");
                printer.Add("repprt");
                printer.Add("locprt");
                printer.Add("micr");
                printer.Add("MX611");

                //create dictionary for printer to port, and printer to driver
                Dictionary<string, string> portTable = new Dictionary<string, string>();
                Dictionary<string, string> driverTable = new Dictionary<string, string>();
                Logger.logging("resetConfig(): Create port and driver dictionary");
                foreach (string pr in printer)
                {
                    if (pr.Equals("MX611", StringComparison.OrdinalIgnoreCase))
                    {
                        portTable.Add(pr, "pBBBBBBm");
                        driverTable.Add(pr, "Lexmark Universal v2 PS3");
                    }
                    else if (pr.Equals("micr", StringComparison.OrdinalIgnoreCase))
                    {
                        portTable.Add(pr, "pBBBBBBb");
                        driverTable.Add(pr, "Lexmark T640 (MS)");
                        //driverTable.Add(pr, "Lexmark 4039 LaserPrinter Plus");
                    }
                    else
                    {
                        portTable.Add(pr, "pBBBBBBa");
                        driverTable.Add(pr, "Lexmark T640 (MS)");
                        //driverTable.Add(pr, "Lexmark 4039 LaserPrinter Plus");
                    }
                }

                //create list of ports
                List<string> port = new List<string>();
                port.Add("pBBBBBBa");
                port.Add("pBBBBBBb");
                port.Add("pBBBBBBg");
                port.Add("pBBBBBBh");
                port.Add("pBBBBBBm");

                //create Dictionary for port to port type
                Dictionary<string, string> portTypeTable = new Dictionary<string, string>();
                Logger.logging("resetConfig(): Create tcp port type [lpr or raw] dictionary");
                foreach (string pr in port)
                {
                    if (pr.Equals("pBBBBBBm", StringComparison.OrdinalIgnoreCase))
                    {
                        portTypeTable.Add(pr, "raw");
                    }
                    else
                    {
                        portTypeTable.Add(pr, "lpr");
                    }
                }


                //create Dictionary for port to port host
                Dictionary<string, string> portHostTable = new Dictionary<string, string>();
                Logger.logging("resetConfig(): Create tcp port host dictionary");
                foreach (string pr in port)
                {
                    //port host default to p000000x
                    portHostTable.Add(pr, pr.Replace('B', '0'));
                }

                bool fail = false;

                //remove printers
                foreach (string pr in printer)
                {
                    Logger.logging("resetConfig(): Check if printer " + pr + " exist before removing");
                    //remove existing printers only
                    if (isPrinterExist(pr))
                    {
                        Logger.logging("resetConfig(): Remove printer " + pr);
                        if (!removePrinter(pr))
                        {
                            Logger.logging("resetConfig(): Unable to remove printer " + pr);
                            fail = true;
                        }
                    }
                }

                //remove ports
                foreach (string pr in port)
                {
                    Logger.logging("resetConfig(): Check if port " + pr + " exist before removing");
                    //remove existing ports only
                    if (isPortExist(pr))
                    {
                        Logger.logging("resetConfig(): Remove port " + pr);
                        if (!removePort(pr))
                        {
                            Logger.logging("resetConfig(): Unable to remove port " + pr);
                            fail = true;
                        }
                    }
                }

                //install port
                foreach (string pr in port)
                {
                    //let's make sure the dictionary has this port before continue
                    Logger.logging("resetConfig(): Check if port type is in dictionary before add port " + pr);
                    if (portTypeTable.ContainsKey(pr) && portHostTable.ContainsKey(pr))
                    {
                        Logger.logging("resetConfig(): Add port " + pr);
                        //addTcpPort(string portName, string portType, string host)
                        if (!addTcpPort(pr, portTypeTable[pr], portHostTable[pr]))
                        {
                            Logger.logging("resetConfig(): Unable to add tcp port " + pr);
                            fail = true;
                        }
                    }
                    else
                    {
                        Logger.logging("resetConfig(): Unable to determine port type and/or port host");
                        fail = true;
                    }
                }

                //install printer
                foreach (string pr in printer)
                {
                    //make sure key exist for port and driver dictionary
                    Logger.logging("resetConfig(): Check if port and driver are in dictionary before add printer " + pr);
                    if (portTable.ContainsKey(pr) && driverTable.ContainsKey(pr))
                    {
                        Logger.logging("resetConfig(): Add printer " + pr);
                        //addPrinter(string printer, string port, string driver)
                        if (!addPrinter(pr, portTable[pr], driverTable[pr]))
                        {
                            Logger.logging("resetConfig(): Unable to add printer " + pr);
                            fail = true;
                        }
                    }
                    else
                    {
                        Logger.logging("resetConfig(): Unable to determine port and/or driver");
                        fail = true;
                    }
                }

                //reset mx611 (printer m) as the default
                Logger.logging("Set printer m (MX611) to default");
                if (setDefaultPrinter("MX611"))
                {
                    Logger.logging("Set default printer m successful");
                }
                else
                {
                    Logger.logging("Failed to set default printer m");
                }

                //print the final message before leaving
                if (fail)
                {
                    printNewLine();
                    printToConsole(" - Failed to reset printer configuration");
                }
                else
                {
                    printNewLine();
                    printToConsole(" - Successfully reset printer configuration");
                }

                printNewLine();
                Program.keyToContinue();
            }
            catch (Exception e)
            {
                Logger.logging("resetConfig() Exception: " + e.Message);
            }
        }
        /// <summary>
        /// Get list of non standard ports - starts with p, follow by any number, and ends with either [a,b,g,h]
        /// </summary>
        /// <returns></returns>
        public static List<string> getNonStandardPort()
        {
            //create list of non-standard port
            List<string> port = new List<string>();

            try
            {
                ProcessStartInfo startInfo = new ProcessStartInfo();
                startInfo.CreateNoWindow = true;
                startInfo.UseShellExecute = false;
                startInfo.RedirectStandardOutput = true;
                startInfo.RedirectStandardError = true;
                startInfo.FileName = getCScriptPath() + "cscript";
                startInfo.Arguments = getPrinterScriptPath() + "prnport.vbs -l";
                //Console.WriteLine("\n" + startInfo.FileName + " " + startInfo.Arguments);
                int exitCode = 1;
                string stdOut = null;
                string stdErr = null;
                string errorAndDescription = null;

                //Start the process with the info we specified
                //Call WaitForExit and then the using statement will close
                using (Process exeProcess = Process.Start(@startInfo))
                {
                    //stdout
                    stdOut = exeProcess.StandardOutput.ReadToEnd();

                    //stderr
                    stdErr = exeProcess.StandardError.ReadToEnd();

                    //get error and description from stdOut and stdErr
                    errorAndDescription = errorAndDescription + getErrorAndDescription(stdOut);
                    errorAndDescription = errorAndDescription + getErrorAndDescription(stdErr);

                    //wait for process to exit
                    exeProcess.WaitForExit();

                    //assign process exitcode
                    //note: exit code of cscript will always return 0 unless cscript is not present
                    //      do not use exit code to check status of printer script
                    //      instead, grep stdOut for error to check for problems with printer script
                    exitCode = exeProcess.ExitCode;

                    if (exitCode != 0)
                    {
                        Logger.logging("getNonStandardPort() error bad exit code: " + exitCode);
                        Logger.logging("getNonStandardPort() error: " + errorAndDescription);
                    }
                    else
                    {
                        //split the output into each line separated by newline
                        string[] lines = stdOut.Split(new string[] { Environment.NewLine }, StringSplitOptions.None);

                        //add the port to list
                        foreach (string line in lines)
                        {
                            if (!string.IsNullOrEmpty(line.Trim()))
                            {
                                try
                                {
                                    if (line.IndexOf("port name", StringComparison.OrdinalIgnoreCase) >= 0)
                                    {
                                        string prt = line.Split(' ')[2].Trim();
                                        if (!string.IsNullOrEmpty(prt))
                                        {
                                            //make sure port name is non-standard before adding
                                            //Logger.logging("Check if port " + prt + " is non-standard port");
                                            if (isNonStandardPort(prt))
                                            {
                                                //Logger.logging("getNonStandardPort(): add non-standard port " + prt + " to list");
                                                port.Add(prt);
                                            }
                                        }
                                    }
                                }
                                catch (Exception e)
                                {
                                    Logger.logging("getNonStandardPort() Exception: " + e.Message);
                                }
                            }
                        }
                    }

                }
            }
            catch (Exception e)
            {
                Logger.logging("getNonStandardPort() Exception: " + e.Message);
            }

            return port;
        }
        /// <summary>
        /// Return true if the port is non-standard
        /// </summary>
        /// <param name="port"></param>
        /// <returns></returns>
        public static bool isNonStandardPort(string port)
        {
            int dummyInt = -1;

            try
            {
                //non-standard port: p[1+ digit][a,b,g,h,m]
                //port length must be 3 char
                //port must start with p
                //port index[1,length - 1] must be a number
                //port must end with [a,b,g,h,m]
                if (!string.IsNullOrEmpty(port) && (port.Length >= 3) && port[0].Equals('p') && (port[port.Length - 1].Equals('a') || port[port.Length - 1].Equals('b') ||
                    port[port.Length - 1].Equals('g') || port[port.Length - 1].Equals('h') ||
                    port[port.Length - 1].Equals('m')) && int.TryParse(port.Substring(1, (port.Length - 2)), out dummyInt))
                {
                    return true;
                }
                else
                {
                    return false;
                }
                
            }
            catch (Exception e)
            {
                Logger.logging("isNonStandardPort() Exception: " + e.Message); 
            }

            return false;
         
        }
        /// <summary>
        /// Return host name/ip of the printer port
        /// </summary>
        /// <returns></returns>
        public static string getPortHost(string port)
        {
            //only proceed if port exists
            if (!string.IsNullOrEmpty(port) && !isPortExist(port))
            {
                return null;
            }

            string portHost = null;

            try
            {
            ProcessStartInfo startInfo = new ProcessStartInfo();
            startInfo.CreateNoWindow = true;
            startInfo.UseShellExecute = false;
            startInfo.RedirectStandardOutput = true;
            startInfo.RedirectStandardError = true;
            startInfo.FileName = getCScriptPath() + "cscript";
            startInfo.Arguments = getPrinterScriptPath() + "prnport.vbs" + " -g -r " + port;
            //Console.WriteLine("\n" + startInfo.FileName + " " + startInfo.Arguments);
            int exitCode = 1;
            string stdOut = null;
            string stdErr = null;
            string errorAndDescription = null;

                //Start the process with the info we specified
                //Call WaitForExit and then the using statement will close
                using (Process exeProcess = Process.Start(@startInfo))
                {
                    //stdout
                    stdOut = exeProcess.StandardOutput.ReadToEnd();

                    //stderr
                    stdErr = exeProcess.StandardError.ReadToEnd();

                    //get error and description from stdOut and stdErr
                    errorAndDescription = errorAndDescription + getErrorAndDescription(stdOut);
                    errorAndDescription = errorAndDescription + getErrorAndDescription(stdErr);

                    //wait for process to exit
                    exeProcess.WaitForExit();

                    //assign process exitcode
                    //note: exit code of cscript will always return 0 unless cscript is not present
                    //      do not use exit code to check status of printer script
                    //      instead, grep stdOut for error to check for problems with printer script
                    exitCode = exeProcess.ExitCode;

                    if (exitCode != 0)
                    {
                        Logger.logging("getPortHost() error bad exit code: " + exitCode);
                        Logger.logging("getPortHost() error: " + errorAndDescription);
                    }
                    else
                    {
                        if (stdOut.IndexOf("error", StringComparison.OrdinalIgnoreCase) >= 0)
                        {
                            Logger.logging("getPortHost() error: " + errorAndDescription);
                        }
                        else if (stdOut.IndexOf("port name", StringComparison.OrdinalIgnoreCase) >= 0)
                        {
                            //split the output into each line separated by newline
                            string[] lines = stdOut.Split(new string[] { Environment.NewLine }, StringSplitOptions.None);

                            string search = "host address";

                            //get the host from string "host address host"
                            foreach (string line in lines)
                            {
                                if (line.IndexOf(search, StringComparison.OrdinalIgnoreCase) >= 0)
                                {
                                    //split "port name host" into array delimited by a space
                                    string[] tmp = line.Split(' ');
                                    portHost = tmp[2].Trim();
                                    break;
                                }
                            }
                        }
                        else
                        {
                            Logger.logging("getPortHost() error: " + errorAndDescription);
                        }
                    }

                }
            }
            catch (Exception e)
            {
                Logger.logging("getPortHost() Exception: " + e.Message);
            }

            return portHost;
        }
        /// <summary>
        /// Return error and description text from message
        /// </summary>
        /// <returns></returns>
        public static string getErrorAndDescription(string message)
        {
            string errorMessage = null;

            try
            {
                string[] lines = message.Split(new string[] { Environment.NewLine }, StringSplitOptions.None);

                foreach (string line in lines)
                {
                    if (line.IndexOf("error", StringComparison.OrdinalIgnoreCase) >= 0)
                    {
                        errorMessage = errorMessage + line + "\n";
                    }
                    else if (line.IndexOf("description", StringComparison.OrdinalIgnoreCase) >= 0)
                    {
                        errorMessage = errorMessage + line;
                    }
                }
            }
            catch (Exception e)
            {
                Logger.logging("getErrorAndDescription() Exception: " + e.Message);
            }
            return errorMessage;
        }
        /// <summary>
        /// Update the printer port host for the port and branch number arguments
        /// </summary>
        /// <param name="port"></param>
        /// <returns></returns>
        public static bool setPortHost(string port, string host)
        {
            //return false if any argument is null
            if (string.IsNullOrEmpty(port) || string.IsNullOrEmpty(host))
            {
                return false;
            }

            //return false if port does not exist or host naming is invalid
            if (!isPortExist(port) && !Program.isHostValidForPort(port, host))
            {
                Logger.logging("setPortHost() port " + port + " does not exist and/or host " + host + " is invalid");
                return false;
            }

            try
            {
                ProcessStartInfo startInfo = new ProcessStartInfo();
                startInfo.CreateNoWindow = true;
                startInfo.UseShellExecute = false;
                startInfo.RedirectStandardOutput = true;
                startInfo.RedirectStandardError = true;
                startInfo.FileName = getCScriptPath() + "cscript";
                //Console.WriteLine("\n" + startInfo.FileName + " " + startInfo.Arguments);
                int exitCode = 1;
                string stdOut = null;
                string stdErr = null;
                string errorAndDescription = null;

                //command switch for tcp raw port - m
                if (string.Equals(port, "pBBBBBBm", StringComparison.OrdinalIgnoreCase))
                {
                    startInfo.Arguments = getPrinterScriptPath() + "prnport.vbs" + " -t -h " + host + " -r " + port;
                }
                else
                {
                    //command switch for other tcp lpr ports - a,b,g,h
                    startInfo.Arguments = getPrinterScriptPath() + "prnport.vbs" + " -t -h " + host + " -q " + host + " -r " + port;
                }

                //Start the process with the info we specified
                //Call WaitForExit and then the using statement will close
                using (Process exeProcess = Process.Start(@startInfo))
                {
                    //stdout
                    stdOut = exeProcess.StandardOutput.ReadToEnd();

                    //stderr
                    stdErr = exeProcess.StandardError.ReadToEnd();

                    //get error and description from stdOut and stdErr
                    errorAndDescription = errorAndDescription + getErrorAndDescription(stdOut);
                    errorAndDescription = errorAndDescription + getErrorAndDescription(stdErr);

                    //wait for process to exit
                    exeProcess.WaitForExit();

                    //assign process exitcode
                    //note: exit code of cscript will always return 0 unless cscript is not present
                    //      do not rely on exit code to check status of printer script
                    //      instead, grep stdOut for error to check for problems with printer script
                    exitCode = exeProcess.ExitCode;

                    if (exitCode != 0)
                    {
                        Logger.logging("setPortHost() error bad exit code: " + exitCode);
                        Logger.logging("setPortHost() error: " + errorAndDescription);
                    }
                    else
                    {
                        if (stdOut.IndexOf("error", StringComparison.OrdinalIgnoreCase) >= 0)
                        {
                            Logger.logging("setPortHost() error: " + errorAndDescription);
                        }
                    }
                }

                //check port existence and return status
                if (string.Equals(getPortHost(port), host))
                {
                    return true;
                }
                else
                {
                    return false;
                }
                
            }
            catch (Exception e)
            {
                Logger.logging("setPortHost() Exception: " + e.Message);
            }

            return false;

        }
        /// <summary>
        /// Update port host for all ports[a,b,g,h,m]
        /// </summary>
        /// <param name="port"></param>
        /// <returns></returns>
        public static void setPortHost(string branch)
        {
            //exit function if branch is null
            if (string.IsNullOrEmpty(branch))
            {
                return;
            }

            try
            {
                //create list of printer port and new host delimited by comma
                List<string> portHost = new List<string>();
                portHost.Add("pBBBBBBa,p" + branch + "a");
                portHost.Add("pBBBBBBb,p" + branch + "b");
                portHost.Add("pBBBBBBg,p" + branch + "g");
                portHost.Add("pBBBBBBh,p" + branch + "h");
                portHost.Add("pBBBBBBm,p" + branch + "m");

                ProcessStartInfo startInfo = new ProcessStartInfo();
                startInfo.CreateNoWindow = true;
                startInfo.UseShellExecute = false;
                startInfo.RedirectStandardOutput = true;
                startInfo.RedirectStandardError = true;
                startInfo.FileName = getCScriptPath() + "cscript";
                //Console.WriteLine("\n" + startInfo.FileName + " " + startInfo.Arguments);
                int exitCode = 1;
                string stdOut = null;
                string stdErr = null;
                string port = null;
                string host = null;
                string errorAndDescription = null;

                foreach (string line in portHost)
                {
                    //split list into port and host
                    port = line.Split(',')[0];
                    host = line.Split(',')[1];

                    //create string of command and assign to startInfo
                    if (string.Equals(port, "pBBBBBBm", StringComparison.OrdinalIgnoreCase))
                    {
                        startInfo.Arguments = getPrinterScriptPath() + "prnport.vbs" + " -t -h " + host + " -r " + port;
                    }
                    else
                    {
                        //also update queue name for lpr ports (a,b,g,h)
                        startInfo.Arguments = getPrinterScriptPath() + "prnport.vbs" + " -t -h " + host + " -q " + host + " -r " + port;
                    }

                    //Start the process with the info we specified
                    //Call WaitForExit and then the using statement will close
                    using (Process exeProcess = Process.Start(@startInfo))
                    {
                        //stdout
                        stdOut = exeProcess.StandardOutput.ReadToEnd();

                        //stderr
                        stdErr = exeProcess.StandardError.ReadToEnd();

                        //get error and description from stdOut and stdErr
                        errorAndDescription = errorAndDescription + getErrorAndDescription(stdOut);
                        errorAndDescription = errorAndDescription + getErrorAndDescription(stdErr);

                        //wait for process to exit
                        exeProcess.WaitForExit();

                        //assign process exitcode
                        //note: exit code of cscript will always return 0 unless cscript is not present
                        //      do not rely on exit code to check status of printer script
                        //      instead, grep stdOut for error to check for problems with printer script
                        exitCode = exeProcess.ExitCode;

                        if (exitCode != 0)
                        {
                            Logger.logging("setPortHost() error bad exit code: " + exitCode);
                            Logger.logging("setPortHost() error: " + errorAndDescription);
                        }
                        else
                        {
                            if (stdOut.IndexOf("error", StringComparison.OrdinalIgnoreCase) >= 0)
                            {
                                Logger.logging("setPortHost() error: " + errorAndDescription);
                            }
                        }
                    }
                }
            }
            catch (Exception e)
            {
                Logger.logging("setPortHost() Exception: " + e.Message);
            }

        }
        /// <summary>
        /// Set default printer
        /// </summary>
        /// <param name="printer"></param>
        /// <returns></returns>
        public static bool setDefaultPrinter(string printer)
        {
            //return false if printer char is null
            if (string.IsNullOrEmpty(printer))
            {
                Logger.logging("setDefaultPrinter() printer arg is null");
                return false;
            }

            try
            {
                //log message and return false if new default printer does not exist
                if (!string.IsNullOrEmpty(printer) && !isPrinterExist(printer))
                {
                    Logger.logging("setDefaultPrinter() unable to set new default printer "
                                    + printer + " because it does not exist");
                    return false;
                }

                //create ConnectionOptions object
                ConnectionOptions connOptions = new ConnectionOptions();

                //create and connect ManagementScope object
                ManagementScope manScope = new ManagementScope(@"\\.\ROOT\CIMV2", connOptions);
                manScope.Connect();

                //create ObjectGetOptions object
                ObjectGetOptions objectGetOptions = new ObjectGetOptions();

                //create ManagementPath object new  process
                ManagementPath managementPath = new ManagementPath("Win32_Printer");

                //hook up everything to the ManagementClass
                ManagementClass printerClass = new ManagementClass(manScope, managementPath, objectGetOptions);

                SelectQuery sq = new SelectQuery();
                sq.QueryString = @"SELECT * FROM Win32_Printer WHERE Name = '" + printer.Replace("\\", "\\\\") + "'";

                ManagementObjectSearcher mos = new ManagementObjectSearcher(manScope, sq);
                ManagementObjectCollection moc = mos.Get();

                //change default printer
                if (moc.Count != 0)
                {
                    foreach (ManagementObject mo in moc)
                    {
                        mo.InvokeMethod("SetDefaultPrinter", new object[] { printer });
                    }
                }

                //check status and return
                string defaultPrinter = getDefaultPrinter();
                if (!string.IsNullOrEmpty(defaultPrinter) && string.Equals(printer, defaultPrinter))
                {
                    return true;
                }
                else
                {
                    return false;
                }
            }
            catch (Exception e)
            {
                //argumentnullexception
                Logger.logging("setDefaultPrinter() Exception: " + e.Message);
            }

            return false;

        }
        /// <summary>
        /// Return list of printers
        /// </summary>
        /// <returns></returns>
        public static List<string> getPrinterList()
        {
            //create a new list of printers
            List<string> printer = new List<string>();

            try
            {

                //define wmi remote connection object
                ConnectionOptions co = new ConnectionOptions();

                ManagementPath mPath = new ManagementPath(@"\\.\root\cimv2");
                ManagementScope mScope = new ManagementScope(mPath, co);

                mScope.Connect();

                ObjectQuery oQuery = new ObjectQuery("SELECT * FROM Win32_Printer");

                ManagementObjectSearcher mObjSearcher = new ManagementObjectSearcher(mScope, oQuery);

                ManagementObjectCollection getObjColl = mObjSearcher.Get();

                //add process to list
                foreach (ManagementObject mObj in getObjColl)
                {
                    printer.Add((string)mObj["Name"]);
                }

            }
            catch (Exception e)
            {
                Logger.logging("getPrinterList() Exception: " + e.Message);
            }

            return printer;
        }
        /// <summary>
        /// Return default printer name
        /// </summary>
        /// <returns></returns>
        public static string getDefaultPrinter()
        {
            string defaultPrinter = null;
            try
            {

                //define wmi remote connection object
                ConnectionOptions co = new ConnectionOptions();

                ManagementPath mPath = new ManagementPath(@"\\.\root\cimv2");
                ManagementScope mScope = new ManagementScope(mPath, co);

                mScope.Connect();

                ObjectQuery oQuery = new ObjectQuery("SELECT * FROM Win32_Printer Where Default = True");

                ManagementObjectSearcher mObjSearcher = new ManagementObjectSearcher(mScope, oQuery);

                ManagementObjectCollection getObjColl = mObjSearcher.Get();

                //add process to list
                foreach (ManagementObject mObj in getObjColl)
                {
                    defaultPrinter = (string)mObj["Name"];
                }

            }
            catch (Exception e)
            {
                Logger.logging("getDefaultPrinter() Exception: " + e.Message);
            }
            return defaultPrinter;
        }
        /// <summary>
        /// Return printer port configured to a printer
        /// </summary>
        /// <returns></returns>
        public static string getPrinterPort(string printer)
        {
            //return null if printer is null
            if (string.IsNullOrEmpty(printer))
            {
                Logger.logging("getPrinterPort() null argument: printer argument is null");
                return null;
            }

            string port = null;
            try
            {

                //define wmi remote connection object
                ConnectionOptions co = new ConnectionOptions();

                ManagementPath mPath = new ManagementPath(@"\\.\root\cimv2");
                ManagementScope mScope = new ManagementScope(mPath, co);

                mScope.Connect();

                ObjectQuery oQuery = new ObjectQuery("SELECT * FROM Win32_Printer Where Name = '" + printer.Replace("\\", "\\\\") + "'");

                ManagementObjectSearcher mObjSearcher = new ManagementObjectSearcher(mScope, oQuery);

                ManagementObjectCollection getObjColl = mObjSearcher.Get();

                //add process to list
                foreach (ManagementObject mObj in getObjColl)
                {
                    port = (string)mObj["PortName"];
                }

            }
            catch (Exception e)
            {
                printNewLine();
                printToConsole("getPrinterPort() Exception: " + e.Message);
            }
            return port;
        }
        /// <summary>
        /// Remove printer port and return true if removal successful
        /// </summary>
        /// <param name="portName"></param>
        /// <returns></returns>
        public static bool removePort(string portName)
        {
            //return false if port name is null
            if (string.IsNullOrEmpty(portName))
            {
                Logger.logging("removePort() null argument: port name is null");
                return false;
            }

            try
            {
                ProcessStartInfo startInfo = new ProcessStartInfo();
                startInfo.CreateNoWindow = true;
                startInfo.UseShellExecute = false;
                startInfo.RedirectStandardOutput = true;
                startInfo.RedirectStandardError = true;
                startInfo.FileName = getCScriptPath() + "cscript";
                startInfo.Arguments = getPrinterScriptPath() + "prnport.vbs -d -r \"" + portName + "\"";

                //Console.WriteLine("\n" + startInfo.FileName + " " + startInfo.Arguments);
                int exitCode = 1;
                string stdOut = null;
                string stdErr = null;
                string errorAndDescription = null;

                //Start the process with the info we specified
                //Call WaitForExit and then the using statement will close
                using (Process exeProcess = Process.Start(@startInfo))
                {
                    //stdout
                    stdOut = exeProcess.StandardOutput.ReadToEnd();

                    //stderr
                    stdErr = exeProcess.StandardError.ReadToEnd();

                    //get error and description from stdOut and stdErr
                    errorAndDescription = errorAndDescription + getErrorAndDescription(stdOut);
                    errorAndDescription = errorAndDescription + getErrorAndDescription(stdErr);

                    //wait for process to exit
                    exeProcess.WaitForExit();

                    //assign process exitcode
                    //note: exit code of cscript will always return 0 unless cscript is not present
                    //      do not use exit code to check status of printer script
                    //      instead, grep stdOut for error to check for problems with printer script
                    exitCode = exeProcess.ExitCode;

                    if (exitCode != 0)
                    {
                        Logger.logging("removePort() error bad exit code: " + exitCode);
                        Logger.logging("removePort() error: " + errorAndDescription);
                    }
                    else
                    {
                        if (!isPortExist(portName))
                        {
                            return true;
                        }
                        else
                        {
                            return false;
                        }
                    }

                }
            }
            catch (Exception e)
            {
                Logger.logging("removePort() Exception: " + e.Message);
            }

            return false;
        }
        /// <summary>
        /// Remove printer and return true if successful
        /// </summary>
        /// <param name="portName"></param>
        /// <returns></returns>
        public static bool removePrinter(string printer)
        {
            //return false if argument is null
            if (string.IsNullOrEmpty(printer))
            {
                Logger.logging("removePrinter() null argument: printer argument is null");
                return false;
            }

            try
            {
                ProcessStartInfo startInfo = new ProcessStartInfo();
                startInfo.CreateNoWindow = true;
                startInfo.UseShellExecute = false;
                startInfo.RedirectStandardOutput = true;
                startInfo.RedirectStandardError = true;
                startInfo.FileName = getCScriptPath() + "cscript";
                startInfo.Arguments = getPrinterScriptPath() + "prnmngr.vbs -d -p \"" + printer + "\"";

                //Console.WriteLine("\n" + startInfo.FileName + " " + startInfo.Arguments);
                int exitCode = 1;
                string stdOut = null;
                string stdErr = null;
                string errorAndDescription = null;

                //Start the process with the info we specified
                //Call WaitForExit and then the using statement will close
                using (Process exeProcess = Process.Start(@startInfo))
                {
                    //stdout
                    stdOut = exeProcess.StandardOutput.ReadToEnd();

                    //stderr
                    stdErr = exeProcess.StandardError.ReadToEnd();

                    //get error and description from stdOut and stdErr
                    errorAndDescription = errorAndDescription + getErrorAndDescription(stdOut);
                    errorAndDescription = errorAndDescription + getErrorAndDescription(stdErr);

                    //wait for process to exit
                    exeProcess.WaitForExit();

                    //assign process exitcode
                    //note: exit code of cscript will always return 0 unless cscript is not present
                    //      do not use exit code to check status of printer script
                    //      instead, grep stdOut for error to check for problems with printer script
                    exitCode = exeProcess.ExitCode;

                    if (exitCode != 0)
                    {
                        Logger.logging("removePrinter() error bad exit code: " + exitCode);
                        Logger.logging("removePrinter() error: " + errorAndDescription);
                    }
                    else
                    {
                        if (!isPrinterExist(printer))
                        {
                            return true;
                        }
                        else
                        {
                            return false;
                        }
                    }

                }
            }
            catch (Exception e)
            {
                Logger.logging("removePrinter() Exception: " + e.Message);
            }

            return false;
        }
        /// <summary>
        /// Add a new printer and return true if successful
        /// </summary>
        /// <param name="printer"></param>
        /// <param name="port"></param>
        /// <param name="driver"></param>
        /// <returns></returns>
        public static bool addPrinter(string printer, string port, string driver)
        {
            //return false if one or more arguments are null
            if (string.IsNullOrEmpty(printer) || string.IsNullOrEmpty(port) ||
                string.IsNullOrEmpty(driver))
            {
                Logger.logging("addPrinter() null argument: one or more arguments is/are null");
                return false;
            }

            try
            {
            ProcessStartInfo startInfo = new ProcessStartInfo();
            startInfo.CreateNoWindow = true;
            startInfo.UseShellExecute = false;
            startInfo.RedirectStandardOutput = true;
            startInfo.RedirectStandardError = true;
            startInfo.FileName = getCScriptPath() + "cscript";
            //cscript prnmngr.vbs -a -p "locprt" -r "pBBBBBBb" -m "Lexmark T640 (MS)"
            startInfo.Arguments = getPrinterScriptPath() + "prnmngr.vbs -a -p \"" + printer + "\" -r \"" +
                                  port + "\" -m \"" + driver + "\"";

            //Console.WriteLine("\n" + startInfo.FileName + " " + startInfo.Arguments);
            int exitCode = 1;
            string stdOut = null;
            string stdErr = null;
            string errorAndDescription = null;

                //Start the process with the info we specified
                //Call WaitForExit and then the using statement will close
                using (Process exeProcess = Process.Start(@startInfo))
                {
                    //stdout
                    stdOut = exeProcess.StandardOutput.ReadToEnd();

                    //stderr
                    stdErr = exeProcess.StandardError.ReadToEnd();

                    //get error and description from stdOut and stdErr
                    errorAndDescription = errorAndDescription + getErrorAndDescription(stdOut);
                    errorAndDescription = errorAndDescription + getErrorAndDescription(stdErr);

                    //wait for process to exit
                    exeProcess.WaitForExit();

                    //assign process exitcode
                    //note: exit code of cscript will always return 0 unless cscript is not present
                    //      do not use exit code to check status of printer script
                    //      instead, grep stdOut for error to check for problems with printer script
                    exitCode = exeProcess.ExitCode;

                    if (exitCode != 0)
                    {
                        Logger.logging("addPrinter() error bad exit code: " + exitCode);
                        Logger.logging("addPrinter() error: " + errorAndDescription);
                    }
                    else
                    {
                        if (stdOut.IndexOf("error", StringComparison.OrdinalIgnoreCase) >= 0)
                        {
                            Logger.logging("addPrinter() error: " + errorAndDescription);
                            return false;
                        }
                        else if (stdOut.IndexOf("usage", StringComparison.OrdinalIgnoreCase) >= 0)
                        {
                            //invalid command argument
                            Logger.logging("addPrinter() error: Invalid argument in prnmngr.vbs");
                            return false;
                        }
                        else if (stdOut.IndexOf("added printer " + printer, StringComparison.OrdinalIgnoreCase) >= 0)
                        {
                            return true;
                        }
                        else
                        {
                            Logger.logging("addPrinter() error: " + errorAndDescription);
                            return false;
                        }
                    }
                }
            }
            catch (Exception e)
            {
                Logger.logging("addPrinter() Exception: " + e.Message);
            }

            return false;
        }
        /// <summary>
        /// Add a tcp port. Port type must be lpr or raw.
        /// </summary>
        /// <param name="portName"></param>
        /// <param name="portType"></param>
        /// <returns></returns>
        public static bool addTcpPort(string portName, string portType, string host)
        {
            //return false if one or more arguments are null
            if (string.IsNullOrEmpty(portName) || string.IsNullOrEmpty(portType) || string.IsNullOrEmpty(host))
            {
                Logger.logging("addTcpPort() null argument: one or more arguments is/are null");
                return false;
            }

            //check for valid portType before continues
            if (!portType.Equals("lpr", StringComparison.OrdinalIgnoreCase) && 
                !portType.Equals("raw", StringComparison.OrdinalIgnoreCase)
                )
            {
                Logger.logging("addTcpPort() portType " + portType + " is invalid (not lpr or raw)");
                return false;
            }

            try
            {

                ProcessStartInfo startInfo = new ProcessStartInfo();
                startInfo.CreateNoWindow = true;
                startInfo.UseShellExecute = false;
                startInfo.RedirectStandardOutput = true;
                startInfo.RedirectStandardError = true;
                startInfo.FileName = getCScriptPath() + "cscript";
                //assign the correct argument for lpr or raw
                if (portType.Equals("lpr", StringComparison.OrdinalIgnoreCase))
                {
                    //cscript prnport.vbs -a -me -o lpr -q "p111111a" -h "p111111a" -r "pBBBBBBa"
                    startInfo.Arguments = getPrinterScriptPath() + "prnport.vbs -a -me -o lpr -q \"" + host +
                                          "\" -h \"" + host + "\" -r \"" + portName + "\"";
                }
                else if (portType.Equals("raw", StringComparison.OrdinalIgnoreCase))
                {
                    //cscript prnport.vbs -a -me -o raw -n 9100 -h "p111111m" -r "pBBBBBBm"
                    startInfo.Arguments = getPrinterScriptPath() + "prnport.vbs -a -me -o raw -n 9100 -h \"" +
                                          host + "\" -r \"" + portName + "\"";
                }

                //Console.WriteLine("\n" + startInfo.FileName + " " + startInfo.Arguments);
                int exitCode = 1;
                string stdOut = null;
                string stdErr = null;
                string errorAndDescription = null;

                //Start the process with the info we specified
                //Call WaitForExit and then the using statement will close
                using (Process exeProcess = Process.Start(@startInfo))
                {
                    //stdout
                    stdOut = exeProcess.StandardOutput.ReadToEnd();

                    //stderr
                    stdErr = exeProcess.StandardError.ReadToEnd();

                    //get error and description from stdOut and stdErr
                    errorAndDescription = errorAndDescription + getErrorAndDescription(stdOut);
                    errorAndDescription = errorAndDescription + getErrorAndDescription(stdErr);

                    //wait for process to exit
                    exeProcess.WaitForExit();

                    //assign process exitcode
                    //note: exit code of cscript will always return 0 unless cscript is not present
                    //      do not use exit code to check status of printer script
                    //      instead, grep stdOut for error to check for problems with printer script
                    exitCode = exeProcess.ExitCode;

                    if (exitCode != 0)
                    {
                        Logger.logging("addTcpPort() error bad exit code: " + exitCode);
                        Logger.logging("addTcpPort() error: " + errorAndDescription);
                    }
                    else
                    {
                        if (stdOut.IndexOf("error", StringComparison.OrdinalIgnoreCase) >= 0)
                        {
                            //invalid port type
                            Logger.logging("addTcpPort() error: " + errorAndDescription);
                            return false;
                        }
                        else if (stdOut.IndexOf("usage", StringComparison.OrdinalIgnoreCase) >= 0)
                        {
                            //invalid command argument
                            Logger.logging("addTcpPort() error: Invalid argument in prnport.vbs");
                            return false;
                        }
                        else if (stdOut.IndexOf("created/updated port " + portName, StringComparison.OrdinalIgnoreCase) >= 0)
                        {
                            return true;
                        }
                        else
                        {
                            Logger.logging("addTcpPort() error: " + errorAndDescription);
                            return false;
                        }
                    }
                }
            }
            catch (Exception e)
            {
                Logger.logging("addTcpPort() Exception: " + e.Message);
            }

            return false;
        }
        /// <summary>
        /// Return true if port exists. Note: do not use WMI since it returns only if port is assigned to printer
        /// </summary>
        /// <returns></returns>
        public static bool isPortExist(string port)
        {
            //return false if post is null
            if (string.IsNullOrEmpty(port))
            {
                return false;
            }

            try
            {
                ProcessStartInfo startInfo = new ProcessStartInfo();
                startInfo.CreateNoWindow = true;
                startInfo.UseShellExecute = false;
                startInfo.RedirectStandardOutput = true;
                startInfo.RedirectStandardError = true;
                startInfo.FileName = getCScriptPath() + "cscript";
                startInfo.Arguments = getPrinterScriptPath() + "prnport.vbs" + " -g -r " + port;
                //Console.WriteLine("\n" + startInfo.FileName + " " + startInfo.Arguments);
                int exitCode = 1;
                string stdOut = null;
                string stdErr = null;
                string errorAndDescription = null;

                //Start the process with the info we specified
                //Call WaitForExit and then the using statement will close
                using (Process exeProcess = Process.Start(@startInfo))
                {
                    //stdout
                    stdOut = exeProcess.StandardOutput.ReadToEnd();

                    //stderr
                    stdErr = exeProcess.StandardError.ReadToEnd();

                    //get error and description from stdOut and stdErr
                    errorAndDescription = errorAndDescription + getErrorAndDescription(stdOut);
                    errorAndDescription = errorAndDescription + getErrorAndDescription(stdErr);

                    //wait for process to exit
                    exeProcess.WaitForExit();

                    //assign process exitcode
                    //note: exit code of cscript will always return 0 unless cscript is not present
                    //      do not use exit code to check status of printer script
                    //      instead, grep stdOut for error to check for problems with printer script
                    exitCode = exeProcess.ExitCode;

                    if (exitCode != 0)
                    {
                        Logger.logging("isPortExist() error bad exit code: " + exitCode);
                        Logger.logging("isPortExist() error: " + errorAndDescription);
                        return false;
                    }
                    else
                    {
                        if (stdOut.IndexOf("error", StringComparison.OrdinalIgnoreCase) >= 0)
                        {
                            //Logger.logging("isPortExist() error: " + errorAndDescription);
                            return false;
                        }
                        else if (stdOut.IndexOf("port name " + port, StringComparison.OrdinalIgnoreCase) >= 0)
                        {
                            return true;
                        }
                        else
                        {
                            //Logger.logging("isPortExist() error: " + errorAndDescription);
                            return false;
                        }
                    }

                }
            }
            catch (Exception e)
            {
                Logger.logging("isPortExist() Exception: " + e.Message);
            }

            return false;
        }
        /// <summary>
        /// Return true if printer exists
        /// </summary>
        /// <returns></returns>
        public static bool isPrinterExist(string printer)
        {
            //return false if printer is null
            if (string.IsNullOrEmpty(printer))
            {
                Logger.logging("isPrinterExist() null argument: printer argument is null");
                return false;
            }

            try
            {

                //define wmi remote connection object
                ConnectionOptions co = new ConnectionOptions();

                ManagementPath mPath = new ManagementPath(@"\\.\root\cimv2");
                ManagementScope mScope = new ManagementScope(mPath, co);

                mScope.Connect();

                ObjectQuery oQuery = new ObjectQuery("SELECT * FROM Win32_Printer Where Name = '" + printer.Replace("\\", "\\\\") + "'");

                ManagementObjectSearcher mObjSearcher = new ManagementObjectSearcher(mScope, oQuery);

                ManagementObjectCollection getObjColl = mObjSearcher.Get();

                if (getObjColl.Count > 0)
                {
                    return true;
                }
                else
                {
                    return false;
                }

            }
            catch (Exception e)
            {
                Logger.logging("isPrinterExist() Exception: " + e.Message);
            }

            return false;
        }
        /// <summary>
        /// Return default printer port
        /// </summary>
        /// <returns></returns>
        public static string getDefaultPort()
        {
            string port = null;
            try
            {

                //define wmi remote connection object
                ConnectionOptions co = new ConnectionOptions();

                ManagementPath mPath = new ManagementPath(@"\\.\root\cimv2");
                ManagementScope mScope = new ManagementScope(mPath, co);

                mScope.Connect();

                ObjectQuery oQuery = new ObjectQuery("SELECT * FROM Win32_Printer Where Default = True");

                ManagementObjectSearcher mObjSearcher = new ManagementObjectSearcher(mScope, oQuery);

                ManagementObjectCollection getObjColl = mObjSearcher.Get();

                //add process to list
                foreach (ManagementObject mObj in getObjColl)
                {
                    port = (string)mObj["PortName"];
                }

            }
            catch (Exception e)
            {
                Logger.logging("getDefaultPort() Exception: " + e.Message);
            }
            return port;
        }
        /// <summary>
        /// Check if printer name is valid
        /// </summary>
        /// <param name="prName"></param>
        /// <returns>bool</returns>
        public static bool isPrinterNameValid(string prName)
        {
            if (!string.IsNullOrEmpty(prName) && (string.Equals("hplaser", prName, StringComparison.Ordinal) ||
                string.Equals("repprt", prName, StringComparison.Ordinal) ||
                string.Equals("micr", prName, StringComparison.Ordinal) ||
                string.Equals("locprt", prName, StringComparison.Ordinal) ||
                string.Equals("MX611", prName, StringComparison.Ordinal))
                )
            {
                return true;
            }
            else
            {
                return false;
            }
        }
        /// <summary>
        /// Print the error message based on exit code
        /// </summary>
        /// <param name="exitCode"></param>
        /// <param name="stdErr"></param>
        public static void printWmicErrorMessage(int exitCode, string stdErr)
        {
            //exit code:stderr code
            //-2147024891:0x80070005 -> access denied
            //-2147023174:0x800706ba -> workstation not reacheable
            //-2147217308:0x80041064 -> user credentials cannot be used for local connections
            printNewLine();
            if (exitCode == -2147024891)
            {
                Logger.logging("printWmicErrorMessage(): Access is denied");
            }
            else if (exitCode == -2147023174)
            {
                Logger.logging("printWmicErrorMessage(): RPC server is unavailable!");
            }
            else if (exitCode == -2147217308)
            {
                Logger.logging("printWmicErrorMessage(): User credentials cannot be used for local connections");
            }
            else if (exitCode == 0)
            {
                //do nothing
            }
            else
            {
                Logger.logging("printWmicErrorMessage(): unable to decode the meaning of exit code " + exitCode);
                Logger.logging("printWmicErrorMessage() error: " + stdErr);
            }
        }
    }

    //Class to log message for debugging
    public class Logger
    {
        public static void logging(string message)
        {
            //set log path to null
            string path = null;

            //set log file name
            string file = "olpa.log";

            //create stream writer
            StreamWriter sw = null;

            try
            {
                //use temp dir as log path
                path = Environment.GetEnvironmentVariable("temp");
            }
            catch (Exception e)
            {
                Console.WriteLine("logging() exception: " + e.Message);

                //unable to get env path so set path to null or current dir
                path = null;
            }

            //append "\\" if path is not null
            if (!string.IsNullOrEmpty(path))
            {
                path = path + "\\";
            }

            //add path to file
            file = path + file;

            //write message to file
            try
            {
                //create object to append content to file
                sw = new StreamWriter(file, true);
                sw.WriteLine(DateTime.Now.ToString("MM/dd/yy HH:mm") + " " + message);
            }
            catch (Exception e)
            {
                Console.WriteLine("logging() exception: " + e.Message);
            }
            finally
            {
                try
                {
                    sw.Close();
                }
                catch (Exception e)
                {
                    Console.WriteLine("logging() exception: " + e.Message);
                }
            }
        }
    }
}

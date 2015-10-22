using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Diagnostics.CodeAnalysis;
using System.DirectoryServices.ActiveDirectory;
using System.Globalization;
using System.Management;
using System.Net;
using System.Net.NetworkInformation;
using System.Runtime.InteropServices;
using System.Security;
using System.Security.Principal;
using System.Text;
using System.Web.Mvc;

namespace McoEasyTool.Controllers
{
    [AllowAnonymous]
    public class ServersController : Controller
    {
        private DataModelContainer db = new DataModelContainer();
        private McoUtilities.UNCAccessWithCredentials UNC_ACCESSOR = new McoUtilities.UNCAccessWithCredentials();

        public static string GetServerOsVersion(string servername)
        {
            ManagementScope scope = new ManagementScope();
            using (McoUtilities.UNC_ACCESS)
            {
                if (McoUtilities.UNC_ACCESS.NetUseWithCredentials(@"\\bcnvdmpos1\BESR1",
                    HomeController.DEFAULT_USERNAME_IMPERSONNATION,
                    HomeController.DEFAULT_DOMAIN_IMPERSONNATION,
                    McoUtilities.Decrypt(HomeController.DEFAULT_PASSWORD_IMPERSONNATION)))
                {
                    try
                    {
                        ConnectionOptions connection = new ConnectionOptions();
                        connection.EnablePrivileges = true;
                        connection.Username = HomeController.DEFAULT_DOMAIN_IMPERSONNATION + "\\" + HomeController.DEFAULT_USERNAME_IMPERSONNATION;
                        connection.Password = McoUtilities.Decrypt(HomeController.DEFAULT_PASSWORD_IMPERSONNATION);

                        string connectionOptions = @"\\" + servername;
                        connectionOptions += @"\root\CIMV2";
                        scope = new ManagementScope(connectionOptions, connection);
                        scope.Connect();

                        SelectQuery query = new SelectQuery("SELECT * FROM Win32_OperatingSystem");
                        ManagementObjectSearcher searcher = new ManagementObjectSearcher(scope, query);
                        string resultat = "";
                        using (ManagementObjectCollection queryCollection = searcher.Get())
                        {
                            foreach (ManagementObject m in queryCollection)
                            {
                                resultat += string.Format("{0}", m["Caption"]) + "&";
                                resultat += string.Format("{0}", m["Version"]);
                            }
                        }
                        return resultat;
                    }
                    catch (UnauthorizedAccessException ex)
                    {
                        return ex.Message;
                    }
                    catch (Exception ex)
                    {
                        return ex.Message;
                    }
                }
                else
                {
                    return "Erreur lors de la connexion à distance.";
                }
            }
        }

        public static string GetVersionName(string osversion, string osname)
        {
            string versionName = "";
            string[] versiontypes = osversion.Split('.');
            if (versiontypes == null)
            {
                return "Inconnue";
            }
            else
            {
                int versionMajor = 0;
                Int32.TryParse(versiontypes[0], out versionMajor);
                int versionMinor = 0;
                Int32.TryParse(versiontypes[1], out versionMinor);
                switch (versionMajor)
                {
                    case 3:
                        versionName = "Windows NT 3";
                        break;
                    case 4:
                        versionName = "Windows NT 4";
                        break;
                    case 5:
                        switch (versionMinor)
                        {
                            case 0:
                                versionName = "Windows 2000";
                                break;
                            case 1:
                                versionName = "Windows XP";
                                break;
                            case 2:
                                if (osname.IndexOf("Windows XP", StringComparison.OrdinalIgnoreCase) > 0)
                                {
                                    versionName = "Windows XP 64-bit";
                                }
                                else
                                {
                                    if (osname.IndexOf("R2", StringComparison.OrdinalIgnoreCase) > 0)
                                    {
                                        versionName = "Windows Server 2003 R2";
                                    }
                                    else
                                    {
                                        versionName = "Windows Server 2003";
                                    }
                                }
                                break;
                            default:
                                versionName = "Inconnue";
                                break;
                        }
                        break;
                    case 6:
                        switch (versionMinor)
                        {
                            case 0:
                                if (osname.IndexOf("Windows Vista", StringComparison.OrdinalIgnoreCase) > 0)
                                {
                                    versionName = "Windows Vista";
                                }
                                else
                                {
                                    versionName = "Windows Server 2008";
                                }
                                break;
                            case 1:
                                if (osname.IndexOf("Windows 7", StringComparison.OrdinalIgnoreCase) > 0)
                                {
                                    versionName = "Windows 7";
                                }
                                else
                                {
                                    versionName = "Windws Server 2008 R2";
                                }
                                break;
                            case 2:
                                if (osname.IndexOf("Windows 8", StringComparison.OrdinalIgnoreCase) > 0)
                                {
                                    versionName = "Windows 8";
                                }
                                else
                                {
                                    versionName = "Windws Server 2012";
                                }
                                break;
                            case 3:
                                if (osname.IndexOf("Windows 8", StringComparison.OrdinalIgnoreCase) > 0)
                                {
                                    versionName = "Windows 8.&";
                                }
                                else
                                {
                                    versionName = "Windws Server 2012 R2";
                                }
                                break;
                            default:
                                versionName = "Inconnue";
                                break;
                        }
                        break;
                    default:
                        versionName = "Inconnue";
                        break;
                }
            }
            return versionName;
        }

        public static string GetOs(string servername)
        {
            string osversion = GetServerOsVersion(servername);
            string[] osinfos = osversion.Split('&');
            if (osinfos.Length == 2)
            {
                return GetVersionName(osinfos[1], osinfos[0]);
            }
            else
            {
                return "Inconnue";
            }
        }

        [HttpPost]
        public string GetDisksList()
        {
            string servername = "";
            try
            {
                servername = Request.Form["servername"];
            }
            catch { }
            if (servername != null && servername.Trim() != "")
            {
                Dictionary<string, string> logical_disks = new Dictionary<string, string>();
                string disks = "";
                var secure = new System.Security.SecureString();
                foreach (char c in McoUtilities.Decrypt(HomeController.SPACE_PASSWORD_IMPERSONNATION))
                {
                    secure.AppendChar(c);
                }
                ConnectionOptions connection = new ConnectionOptions();
                connection.Username = HomeController.SPACE_DOMAIN_AND_USERNAME_IMPERSONNATION;
                connection.SecurePassword = secure;
                connection.EnablePrivileges = true;
                try
                {
                    string strNameSpace = @"\\";
                    if (servername != "")
                        strNameSpace = (servername.IndexOf(@"\\") == -1) ? strNameSpace + servername : servername;
                    else
                        strNameSpace += ".";
                    strNameSpace += @"\root\cimv2";
                    System.Management.ManagementScope managementScope = new System.Management.ManagementScope(strNameSpace, connection);
                    System.Management.ObjectQuery query = new System.Management.ObjectQuery("select * from Win32_LogicalDisk where DriveType=3");
                    ManagementObjectSearcher moSearcher = new ManagementObjectSearcher(managementScope, query);
                    ManagementObjectCollection moCollection = moSearcher.Get();
                    foreach (ManagementObject oReturn in moCollection)
                    {
                        string id = oReturn["DeviceID"].ToString().ToUpper();
                        string info = oReturn["Name"].ToString().Substring(0, 1).ToUpper();
                        if (!logical_disks.ContainsKey(id) &&
                            !logical_disks.ContainsValue(info))
                        {
                            logical_disks.Add(id, info);
                        }
                        //foreach (PropertyData prop in oReturn.Properties)
                        //{
                        //    Console.WriteLine(prop.Name + " " + prop.Value);
                        //}
                        //disks += oReturn["Name"].ToString().Substring(0, 1) + "; ";
                        /*Console.WriteLine("Drive {0}", oReturn["Name"].ToString());
                        Console.WriteLine("  Volume label: {0}", oReturn["VolumeName"].ToString());
                        Console.WriteLine("  File system: {0}", oReturn["FileSystem"].ToString());
                        Console.WriteLine("  Available space to current user:{0, 15} bytes", oReturn["FreeSpace"].ToString());
                        Console.WriteLine("  Total size of drive:            {0, 15} bytes ", oReturn["Size"].ToString());*/
                    }
                    Dictionary<string, string> mapped_disks = GetMappedDisksList(servername);
                    foreach (KeyValuePair<string, string> mapped_disk in mapped_disks)
                    {
                        if (!logical_disks.ContainsKey(mapped_disk.Key.ToUpper()) &&
                            !logical_disks.ContainsValue(mapped_disk.Value.ToUpper()))
                        {
                            logical_disks.Add(mapped_disk.Key.ToUpper(), mapped_disk.Value.ToUpper());
                        }
                    }
                    foreach (KeyValuePair<string, string> logical_disk in logical_disks)
                    {
                        disks += logical_disk.Value + "; ";
                    }
                    if (disks.Length > 2)
                    {
                        disks = disks.Substring(0, disks.Length - 2);
                    }
                }
                catch (Exception ex)
                {
                    Console.WriteLine(ex.Message);
                }
                return disks;
            }
            return null;
        }

        [HttpPost]
        public string GetSharesList()
        {
            string servername = "";
            try
            {
                servername = Request.Form["servername"];
            }
            catch { }
            if (servername != null && servername.Trim().Length > 0)
            {
                List<string> shares_list = new List<string>();
                string list = "";
                ShareCollection shares;
                try
                {
                    shares = ShareCollection.GetShares(servername);
                    foreach (Share share in shares)
                    {
                        string info = share.ToString().ToUpper().Substring(2);
                        if (!shares_list.Contains(info))
                        {
                            shares_list.Add(info);
                        }
                    }
                }
                catch { }
                Dictionary<string, string> mapped_disks = GetMappedDisksList(servername);
                foreach (KeyValuePair<string, string> mapped_disk in mapped_disks)
                {
                    if ((!shares_list.Contains(mapped_disk.Key.ToUpper())) &&
                        (!shares_list.Contains(mapped_disk.Value.ToUpper()))
                       )
                    {
                        shares_list.Add(mapped_disk.Value.ToUpper());
                    }
                }
                foreach (string share in shares_list)
                {
                    list += share + "; ";
                }
                if (list.Length > 2)
                {
                    list = list.Substring(0, list.Length - 2);
                    return list;
                }
            }
            return null;
        }

        public Dictionary<string, string> GetMappedDisksList(string servername)
        {
            Dictionary<string, string> mapped_disks = new Dictionary<string, string>();
            if (servername != null && servername.Trim() != "")
            {
                var secure = new System.Security.SecureString();
                foreach (char c in McoUtilities.Decrypt(HomeController.SPACE_PASSWORD_IMPERSONNATION))
                {
                    secure.AppendChar(c);
                }
                ConnectionOptions connection = new ConnectionOptions();
                connection.Username = HomeController.SPACE_DOMAIN_AND_USERNAME_IMPERSONNATION;
                connection.SecurePassword = secure;
                connection.EnablePrivileges = true;
                try
                {
                    var searcher = new ManagementObjectSearcher(
                        "root\\CIMV2",
                        "SELECT * FROM Win32_MappedLogicalDisk");
                    string strNameSpace = @"\\";
                    if (servername != "")
                        strNameSpace = (servername.IndexOf(@"\\") == -1) ? strNameSpace + servername : servername;
                    else
                        strNameSpace += ".";
                    strNameSpace += @"\root\cimv2";
                    System.Management.ManagementScope managementScope = new System.Management.ManagementScope(strNameSpace, connection);
                    System.Management.ObjectQuery query = new System.Management.ObjectQuery("SELECT * FROM Win32_MappedLogicalDisk");
                    ManagementObjectSearcher moSearcher = new ManagementObjectSearcher(managementScope, query);
                    ManagementObjectCollection moCollection = moSearcher.Get();

                    foreach (ManagementObject queryObj in moCollection)
                    {
                        string id = queryObj["DeviceID"].ToString().ToUpper();
                        string info = (queryObj["ProviderName"].ToString().Length < 4) ?
                                queryObj["Name"].ToString().Substring(0, 1).ToUpper() :
                                queryObj["ProviderName"].ToString().ToUpper();
                        if (!mapped_disks.ContainsKey(id) && !mapped_disks.ContainsValue(info))
                        {
                            mapped_disks.Add(id, info);
                        }
                        /*Console.WriteLine("-----------------------------------");
                        Console.WriteLine("Win32_MappedLogicalDisk instance");
                        Console.WriteLine("-----------------------------------");
                        Console.WriteLine("Access: {0}", queryObj["Access"]);
                        Console.WriteLine("Availability: {0}", queryObj["Availability"]);
                        Console.WriteLine("BlockSize: {0}", queryObj["BlockSize"]);
                        Console.WriteLine("Caption: {0}", queryObj["Caption"]);
                        Console.WriteLine("Compressed: {0}", queryObj["Compressed"]);
                        Console.WriteLine("ConfigManagerErrorCode: {0}", queryObj["ConfigManagerErrorCode"]);
                        Console.WriteLine("ConfigManagerUserConfig: {0}", queryObj["ConfigManagerUserConfig"]);
                        Console.WriteLine("CreationClassName: {0}", queryObj["CreationClassName"]);
                        Console.WriteLine("Description: {0}", queryObj["Description"]);
                        Console.WriteLine("DeviceID: {0}", queryObj["DeviceID"]);
                        Console.WriteLine("ErrorCleared: {0}", queryObj["ErrorCleared"]);
                        Console.WriteLine("ErrorDescription: {0}", queryObj["ErrorDescription"]);
                        Console.WriteLine("ErrorMethodology: {0}", queryObj["ErrorMethodology"]);
                        Console.WriteLine("FileSystem: {0}", queryObj["FileSystem"]);
                        Console.WriteLine("FreeSpace: {0}", queryObj["FreeSpace"]);
                        Console.WriteLine("InstallDate: {0}", queryObj["InstallDate"]);
                        Console.WriteLine("LastErrorCode: {0}", queryObj["LastErrorCode"]);
                        Console.WriteLine("MaximumComponentLength: {0}", queryObj["MaximumComponentLength"]);
                        Console.WriteLine("Name: {0}", queryObj["Name"]);
                        Console.WriteLine("NumberOfBlocks: {0}", queryObj["NumberOfBlocks"]);
                        Console.WriteLine("PNPDeviceID: {0}", queryObj["PNPDeviceID"]);

                        if (queryObj["PowerManagementCapabilities"] == null)
                            Console.WriteLine("PowerManagementCapabilities: {0}", queryObj["PowerManagementCapabilities"]);
                        else
                        {
                            UInt16[] arrPowerManagementCapabilities = (UInt16[])(queryObj["PowerManagementCapabilities"]);
                            foreach (UInt16 arrValue in arrPowerManagementCapabilities)
                            {
                                Console.WriteLine("PowerManagementCapabilities: {0}", arrValue);
                            }
                        }
                        Console.WriteLine("PowerManagementSupported: {0}", queryObj["PowerManagementSupported"]);
                        Console.WriteLine("ProviderName: {0}", queryObj["ProviderName"]);
                        Console.WriteLine("Purpose: {0}", queryObj["Purpose"]);
                        Console.WriteLine("QuotasDisabled: {0}", queryObj["QuotasDisabled"]);
                        Console.WriteLine("QuotasIncomplete: {0}", queryObj["QuotasIncomplete"]);
                        Console.WriteLine("QuotasRebuilding: {0}", queryObj["QuotasRebuilding"]);
                        Console.WriteLine("SessionID: {0}", queryObj["SessionID"]);
                        Console.WriteLine("Size: {0}", queryObj["Size"]);
                        Console.WriteLine("Status: {0}", queryObj["Status"]);
                        Console.WriteLine("StatusInfo: {0}", queryObj["StatusInfo"]);
                        Console.WriteLine("SupportsDiskQuotas: {0}", queryObj["SupportsDiskQuotas"]);
                        Console.WriteLine("SupportsFileBasedCompression: {0}", queryObj["SupportsFileBasedCompression"]);
                        Console.WriteLine("SystemCreationClassName: {0}", queryObj["SystemCreationClassName"]);
                        Console.WriteLine("SystemName: {0}", queryObj["SystemName"]);
                        Console.WriteLine("VolumeName: {0}", queryObj["VolumeName"]);
                        Console.WriteLine("VolumeSerialNumber: {0}", queryObj["VolumeSerialNumber"]);*/
                    }
                }
                catch (ManagementException ex)
                {
                    string message = "An error occurred while querying for WMI data: " + ex.Message;
                }
                return mapped_disks;
            }
            return null;
        }

        public static List<VirtualizedPartition> GetRemainingSpaceOnDisks(string servername, Account account = null)
        {
            List<VirtualizedPartition> partitions = new List<VirtualizedPartition>();
            if (servername != null && servername.Trim() != "")
            {
                var secure = new System.Security.SecureString();
                string password = (account == null) ? McoUtilities.Decrypt(HomeController.SPACE_PASSWORD_IMPERSONNATION) :
                    McoUtilities.Decrypt(account.Password);
                foreach (char c in password)
                {
                    secure.AppendChar(c);
                }
                ConnectionOptions connection = new ConnectionOptions();
                connection.Username = (account == null) ? HomeController.SPACE_DOMAIN_AND_USERNAME_IMPERSONNATION : account.DisplayName;
                connection.SecurePassword = secure;
                connection.EnablePrivileges = true;
                try
                {
                    string strNameSpace = @"\\";
                    if (servername != "")
                        strNameSpace = (servername.IndexOf(@"\\") == -1) ? strNameSpace + servername : servername;
                    else
                        strNameSpace += ".";
                    strNameSpace += @"\root\cimv2";
                    System.Management.ManagementScope managementScope = new System.Management.ManagementScope(strNameSpace, connection);
                    System.Management.ObjectQuery query = new System.Management.ObjectQuery("select * from Win32_LogicalDisk where DriveType=3");
                    ManagementObjectSearcher moSearcher = new ManagementObjectSearcher(managementScope, query);
                    ManagementObjectCollection moCollection = moSearcher.Get();
                    foreach (ManagementObject oReturn in moCollection)
                    {
                        string info = oReturn["Name"].ToString().Substring(0, 1).ToUpper();
                        VirtualizedPartition partition = new VirtualizedPartition(servername, info);
                        partition.Name = partition.Name.Substring(0, 1);
                        partition.Volume = oReturn["VolumeName"].ToString();
                        partition.FileSystem = oReturn["FileSystem"].ToString();
                        partition.AvailableSpace = oReturn["FreeSpace"].ToString() + " " + HomeController.DEFAULT_OCTECT_UNIT;
                        partition.TotalSize = oReturn["Size"].ToString() + " " + HomeController.DEFAULT_OCTECT_UNIT;
                        partitions.Add(partition);
                    }
                }
                catch (Exception ex)
                {
                    Console.WriteLine(ex.Message);
                }
                return partitions;
            }
            return null;
        }

        public static List<VirtualizedPartition> GetRemainingSpaceOnShares(string servername, Account account = null)
        {
            List<VirtualizedPartition> partitions = new List<VirtualizedPartition>();
            if (servername != null && servername.Trim().Length > 0)
            {
                string domain = (account == null) ? HomeController.SPACE_DOMAIN_IMPERSONNATION : account.Domain;
                string username = (account == null) ? HomeController.SPACE_USERNAME_IMPERSONNATION : account.Username;
                string password = (account == null) ? HomeController.SPACE_PASSWORD_IMPERSONNATION : account.Password;

                IntPtr userToken = IntPtr.Zero;
                bool success = McoUtilities.LogonUser(
                  username, domain,
                  McoUtilities.Decrypt(password),
                  (int)McoUtilities.LogonType.LOGON32_LOGON_INTERACTIVE, //2
                  (int)McoUtilities.LogonProvider.LOGON32_PROVIDER_DEFAULT, //0
                  out userToken);
                if (!success)
                {
                    return null;
                }

                using (WindowsIdentity.Impersonate(userToken))
                {
                    ShareCollection shares;
                    try
                    {
                        shares = ShareCollection.GetShares(servername);
                        foreach (Share share in shares)
                        {
                            VirtualizedPartition partition = new VirtualizedPartition(servername, share.ToString());


                            long free = 0, dummy1 = 0, dummy2 = 0;
                            GetDiskFreeSpaceEx(share.ToString(), ref free, ref dummy1, ref dummy2);
                            partition.Volume = dummy2.ToString();
                            //partition.FileSystem = oReturn["FileSystem"].ToString();
                            partition.AvailableSpace = free.ToString() + " " + HomeController.DEFAULT_OCTECT_UNIT;
                            partition.TotalSize = dummy1.ToString() + " " + HomeController.DEFAULT_OCTECT_UNIT;
                            partitions.Add(partition);
                        }
                    }
                    catch { }
                }

                return partitions;
            }
            return null;
        }

        public static List<VirtualizedPartition> GetRemainingSpaceOnMappedDisk(string servername, Account account = null)
        {
            List<VirtualizedPartition> partitions = new List<VirtualizedPartition>();
            if (servername != null && servername.Trim() != "")
            {
                var secure = new System.Security.SecureString();
                string password = (account == null) ? McoUtilities.Decrypt(HomeController.SPACE_PASSWORD_IMPERSONNATION) :
                    McoUtilities.Decrypt(account.Password);
                foreach (char c in password)
                {
                    secure.AppendChar(c);
                }
                ConnectionOptions connection = new ConnectionOptions();
                connection.Username = (account == null) ? HomeController.SPACE_DOMAIN_AND_USERNAME_IMPERSONNATION : account.DisplayName;
                connection.SecurePassword = secure;
                connection.EnablePrivileges = true;
                try
                {
                    string strNameSpace = @"\\";
                    if (servername != "")
                        strNameSpace = (servername.IndexOf(@"\\") == -1) ? strNameSpace + servername : servername;
                    else
                        strNameSpace += ".";
                    strNameSpace += @"\root\cimv2";
                    System.Management.ManagementScope managementScope = new System.Management.ManagementScope(strNameSpace, connection);
                    System.Management.ObjectQuery query = new System.Management.ObjectQuery("SELECT * FROM Win32_MappedLogicalDisk");
                    ManagementObjectSearcher moSearcher = new ManagementObjectSearcher(managementScope, query);
                    ManagementObjectCollection moCollection = moSearcher.Get();

                    foreach (ManagementObject queryObj in moCollection)
                    {
                        string id = queryObj["DeviceID"].ToString().ToUpper();
                        string info = (queryObj["ProviderName"].ToString().Length < 4) ?
                                queryObj["Name"].ToString().Substring(0, 1).ToUpper() :
                                queryObj["ProviderName"].ToString().ToUpper();

                        VirtualizedPartition partition = new VirtualizedPartition(servername, info);
                        partition.Volume = queryObj["VolumeName"].ToString();
                        partition.FileSystem = queryObj["FileSystem"].ToString();
                        partition.AvailableSpace = queryObj["FreeSpace"].ToString() + " " + HomeController.DEFAULT_OCTECT_UNIT;
                        partition.TotalSize = queryObj["Size"].ToString() + " " + HomeController.DEFAULT_OCTECT_UNIT;
                        partitions.Add(partition);
                    }
                }
                catch (ManagementException ex)
                {
                    string message = "An error occurred while querying for WMI data: " + ex.Message;
                }
                return partitions;
            }
            return null;
        }

        public static List<VirtualizedPartition> GetRemainingSpaceOnMappedPartitions(SpaceServer server, Account account = null)
        {
            List<VirtualizedPartition> partitions = new List<VirtualizedPartition>();
            Dictionary<string, string> server_disks = new Dictionary<string, string>();
            string[] parser = server.Disks.Split(new string[] { "; " }, StringSplitOptions.RemoveEmptyEntries);
            for (int index = 0; index < parser.Length; index++)
            {
                string[] infos = parser[index].Split(new string[] { "|" }, StringSplitOptions.RemoveEmptyEntries);
                string disk = infos[0].Substring(infos[0].IndexOf("=") + 1);
                string threshold = infos[1].Substring(infos[1].IndexOf("=") + 1);
                server_disks.Add(disk, threshold);

                Process connect_drive_process = new Process();
                connect_drive_process.StartInfo.FileName = "cmd.exe";
                connect_drive_process.StartInfo.Arguments = "/c net use " + HomeController.SPACE_DEFAULT_LOCAL_MAPPED_DRIVE_LETTER +
                    ": " + disk + " /user:" + HomeController.SPACE_DOMAIN_AND_USERNAME_IMPERSONNATION +
                    " \"" + McoUtilities.Decrypt(HomeController.SPACE_PASSWORD_IMPERSONNATION) + "\"";
                connect_drive_process.Start();
                connect_drive_process.WaitForExit();

                var secure = new System.Security.SecureString();
                foreach (char c in McoUtilities.Decrypt(HomeController.SPACE_PASSWORD_IMPERSONNATION))
                {
                    secure.AppendChar(c);
                }
                ConnectionOptions connection = new ConnectionOptions();
                connection.Username = HomeController.SPACE_DOMAIN_AND_USERNAME_IMPERSONNATION;
                connection.SecurePassword = secure;
                connection.EnablePrivileges = true;
                try
                {
                    var searcher = new ManagementObjectSearcher(
                        "root\\CIMV2",
                        "SELECT * FROM Win32_MappedLogicalDisk");
                    string strNameSpace = @"\\";
                    strNameSpace += ".";
                    strNameSpace += @"\root\cimv2";
                    System.Management.ManagementScope managementScope = new System.Management.ManagementScope(strNameSpace, connection);
                    ManagementObjectCollection moCollection = searcher.Get();

                    foreach (ManagementObject queryObj in moCollection)
                    {
                        string id = queryObj["DeviceID"].ToString().ToUpper();
                        string info = disk;

                        VirtualizedPartition partition = new VirtualizedPartition(server.Name, info);
                        partition.Volume = queryObj["VolumeName"].ToString();
                        partition.FileSystem = queryObj["FileSystem"].ToString();
                        partition.AvailableSpace = queryObj["FreeSpace"].ToString() + " " + HomeController.DEFAULT_OCTECT_UNIT;
                        partition.TotalSize = queryObj["Size"].ToString() + " " + HomeController.DEFAULT_OCTECT_UNIT;
                        partitions.Add(partition);
                    }
                }
                catch { }

                Process disconnect_drive_process = new Process();
                disconnect_drive_process.StartInfo.FileName = "cmd.exe";
                disconnect_drive_process.StartInfo.Arguments = "/c net use " +
                    HomeController.SPACE_DEFAULT_LOCAL_MAPPED_DRIVE_LETTER + ": /delete";
                disconnect_drive_process.Start();
                disconnect_drive_process.WaitForExit();
            }
            return partitions;
        }

        public static VirtualizedPartition GetVirtualizedPartition(List<VirtualizedPartition> partitions, string disk)
        {
            if (disk != null && disk.Trim() != "")
            {
                foreach (VirtualizedPartition partition in partitions)
                {
                    if (partition.Name.ToUpper() == disk.Trim().ToUpper())
                    {
                        return partition;
                    }
                }
            }
            return null;
        }

        public class VirtualizedServer_Result
        {
            public Server COMMON_Server { get; set; }
            public FaultyServer AD_Server { get; set; }
            public BackupServer BESR_Server { get; set; }
            public AppServer APP_Server { get; set; }
            public SpaceServer SPACE_Server { get; set; }
            public string Ping { get; set; }

            public VirtualizedServer_Result()
            {
                this.COMMON_Server = new Server();
                this.Ping = "";
            }

            public VirtualizedServer_Result(Server server)
                : this()
            {
                this.COMMON_Server = server;
            }
        }

        public class VirtualizedServer
        {
            public int Id { get; set; }
            public string IpAddress { get; set; }
            public string Name { get; set; }
            public string Location { get; set; }
            public string Ping { get; set; }
            public string State { get; set; }
            public string Site { get; set; }
            public string IdSite { get; set; }
            public string ActiveDirectoryDomain { get; set; }

            public VirtualizedServer()
            {
                this.Id = 0; this.IpAddress = "";
                this.Name = this.Location = this.ActiveDirectoryDomain = "";
                this.Ping = this.State = this.Site = this.IdSite = "";
            }

            public VirtualizedServer(string Name)
                : this()
            {
                this.Name = Name;
            }

            public VirtualizedServer GetMatchingServer(Dictionary<int, VirtualizedServer> FOREST)
            {
                for (int forestindex = 0; forestindex < FOREST.Count; forestindex++)
                {
                    if ((this.Name != null && this.Name.Trim() != "") &&
                        ((this.Name.ToUpper() == FOREST[forestindex].Name.ToUpper()) ||
                        (this.Name.ToLower() == FOREST[forestindex].Name.ToLower()))
                    )
                    {
                        return FOREST[forestindex];
                    }
                }
                return this;
            }

            public ReftechServers GetMatchingReftechServer(ReftechServers[] REFTECH_SERVERS)
            {
                try
                {
                    foreach (ReftechServers reftechserver in REFTECH_SERVERS)
                    {
                        if (reftechserver.NomMachineServeur.ToUpper() == this.Name.Trim().ToUpper())
                        {
                            return reftechserver;
                        }
                    }
                }
                catch { }
                return null;
            }

            public void TryPing()
            {
                try
                {
                    Ping ping = new Ping();
                    PingOptions options = new PingOptions(64, true);
                    PingReply pingreply = ping.Send(this.Name);
                    this.Ping = (pingreply.Status.ToString() == "Success") ? "OK" : "KO";
                    if (this.IpAddress == null || this.IpAddress.Trim() == "")
                    {
                        this.IpAddress = pingreply.Address.ToString();
                    }
                }
                catch
                {
                    this.Ping = "KO";
                }
            }

            public void SetSite()
            {
                string PathDirectory = HomeController.AD_RESULTS_FOLDER;
                try
                {
                    Process process = new Process();
                    process.StartInfo.FileName = HomeController.BATCHES_FOLDER + "GetSiteName.bat";
                    process.StartInfo.UseShellExecute = false;
                    process.StartInfo.RedirectStandardOutput = false;
                    process.StartInfo.WorkingDirectory = PathDirectory;
                    process.StartInfo.CreateNoWindow = true;
                    process.StartInfo.Arguments = this.ActiveDirectoryDomain;
                    process.Start();
                    process.WaitForExit();
                }
                catch { }

                try
                {
                    string[] lines = System.IO.File.ReadAllLines(PathDirectory + this.ActiveDirectoryDomain + "sitename.txt", Encoding.Default);
                    foreach (string siteline in lines)
                    {
                        if (siteline.IndexOf(this.Name, StringComparison.OrdinalIgnoreCase) != -1)
                        {
                            string selectedline = siteline.Substring(
                                siteline.IndexOf(this.Name, StringComparison.OrdinalIgnoreCase));
                            string[] sitelineParser = selectedline.Split(':');
                            if (sitelineParser.Length > 1)
                            {
                                this.Site = sitelineParser[1].Trim();
                                break;
                            }
                        }
                    }
                    if (this.Site.Trim() == "")
                    {
                        this.Site = "N/A";
                    }
                }
                catch { }
            }

            public void SetInformations(Dictionary<int, VirtualizedServer> FOREST, ReftechServers[] REFTECH_SERVERS, bool ping = false)
            {
                VirtualizedServer server = GetMatchingServer(FOREST);
                this.IdSite = server.IdSite; this.IpAddress = server.IpAddress;
                this.Location = server.Location; this.Site = server.Site;
                ReftechServers reftech_server = GetMatchingReftechServer(REFTECH_SERVERS);
                if (reftech_server != null)
                {
                    this.State = reftech_server.EtatServeur;
                    this.IdSite = reftech_server.IdSite;
                    this.IpAddress = reftech_server.IP;
                    switch (this.State)
                    {
                        case "O": this.State = "Operationnel"; break;
                        case "A": this.State = "A Venir"; break;
                        case "R": this.State = "Retiré"; break;
                        default: this.State = (this.State != null && this.State.Trim() != "") ? this.State : "Inconnu"; break;
                    }
                }
                if (ping)
                {
                    TryPing();
                }
            }
        }

        public class VirtualizedPartition
        {
            public string Owner { get; set; }
            public string Name { get; set; }
            public string Volume { get; set; }
            public string FileSystem { get; set; }
            public string AvailableSpace { get; set; }
            public string TotalSize { get; set; }
            public string Threshold { get; set; }
            public bool Critical { get; set; }

            public VirtualizedPartition(string owner, string name)
            {
                this.Owner = owner;
                this.Name = name;
            }

            public bool IsCritical()
            {
                double threshold = (this.Threshold != null && this.Threshold.Trim() != ""
                    && this.Threshold.IndexOf("o") != -1) ? GetSizeValue(this.Threshold) :
                    HomeController.SPACE_DEFAULT_THRESHOLD;

                double available = GetSizeValue(this.AvailableSpace);
                this.Critical = (available <= threshold);
                this.SwitchToTera();
                return (available <= threshold);
            }

            public void SwitchTo(string to)
            {
                double available = GetSizeValue(this.AvailableSpace);
                available = SizeConversion(available, GetSizeUnit(this.AvailableSpace), to);
                this.AvailableSpace = available.ToString() + " " + to;

                double totalsize = GetSizeValue(this.TotalSize);
                totalsize = SizeConversion(totalsize, GetSizeUnit(this.TotalSize), to);
                this.TotalSize = totalsize.ToString() + " " + to;

                double threshold = GetSizeValue(this.Threshold);
                threshold = SizeConversion(threshold, GetSizeUnit(this.AvailableSpace), to);
                this.Threshold = threshold.ToString() + " " + to;
            }

            public void SwitchToTera()
            {
                this.SwitchTo(HomeController.DEFAULT_TERA_OCTECT_UNIT);
            }

            public void SwitchToGiga()
            {
                this.SwitchTo(HomeController.DEFAULT_GIGA_OCTECT_UNIT);
            }

            public void SwitchToMega()
            {
                this.SwitchTo(HomeController.DEFAULT_MEGA_OCTECT_UNIT);
            }

            public void SwitchToKilo()
            {
                this.SwitchTo(HomeController.DEFAULT_KILO_OCTECT_UNIT);
            }

            public void SwitchToOct()
            {
                this.SwitchTo(HomeController.DEFAULT_OCTECT_UNIT);
            }

            public static double SizeConversion(double size, string from, string to)
            {
                int uppow = 0;
                int downpow = 0;
                switch (from)
                {
                    case HomeController.DEFAULT_TERA_OCTECT_UNIT: uppow = 4; break;
                    case HomeController.DEFAULT_GIGA_OCTECT_UNIT: uppow = 3; break;
                    case HomeController.DEFAULT_MEGA_OCTECT_UNIT: uppow = 2; break;
                    case HomeController.DEFAULT_KILO_OCTECT_UNIT: uppow = 1; break;
                    case HomeController.DEFAULT_OCTECT_UNIT: uppow = 0; break;
                }
                switch (to)
                {
                    case HomeController.DEFAULT_TERA_OCTECT_UNIT: downpow = 4; break;
                    case HomeController.DEFAULT_GIGA_OCTECT_UNIT: downpow = 3; break;
                    case HomeController.DEFAULT_MEGA_OCTECT_UNIT: downpow = 2; break;
                    case HomeController.DEFAULT_KILO_OCTECT_UNIT: downpow = 1; break;
                    case HomeController.DEFAULT_OCTECT_UNIT: downpow = 0; break;

                }
                int pow = uppow - downpow;
                return Math.Round((size * Math.Pow(HomeController.DEFAULT_OCTECT_INCREMENT, pow)), 2);
            }

            public static string GetSizeUnit(string size)
            {
                if (size != null && size.Trim() != "" && size.Split(' ').Length == 2)
                {
                    string unit = size.Split(' ')[1].Trim();
                    switch (unit)
                    {
                        case HomeController.DEFAULT_TERA_OCTECT_UNIT: return HomeController.DEFAULT_TERA_OCTECT_UNIT;
                        case HomeController.DEFAULT_GIGA_OCTECT_UNIT: return HomeController.DEFAULT_GIGA_OCTECT_UNIT;
                        case HomeController.DEFAULT_MEGA_OCTECT_UNIT: return HomeController.DEFAULT_MEGA_OCTECT_UNIT;
                        case HomeController.DEFAULT_KILO_OCTECT_UNIT: return HomeController.DEFAULT_KILO_OCTECT_UNIT;
                        case HomeController.DEFAULT_OCTECT_UNIT: return HomeController.DEFAULT_OCTECT_UNIT;
                    }
                    return unit;
                }
                return null;
            }

            public static double GetSizeValue(string size)
            {
                if (size != null && size.Trim() != "" && size.Split(' ').Length == 2)
                {
                    NumberFormatInfo provider = new NumberFormatInfo();
                    provider.NumberDecimalSeparator = HomeController.DEFAULT_NUMBER_DECIMAL_SEPARATOR;
                    string value = size.Split(' ')[0];
                    return Convert.ToDouble(value, provider);
                }
                return 0;
            }

            public static bool HasUnit(string size, string unit)
            {
                string current_unit = GetSizeUnit(size);
                if (current_unit == unit)
                {
                    return true;
                }
                return false;
            }
        }

        public static VirtualizedServer_Result GetServerInformations(Dictionary<int, VirtualizedServer> FOREST, ReftechServers[] REFTECH_SERVERS, Server server, string module, bool ping = false)
        {
            VirtualizedServer virtual_server = new VirtualizedServer(server.Name);
            virtual_server.SetInformations(FOREST, REFTECH_SERVERS, ping);
            server.Location = virtual_server.Location; server.IpAddress = virtual_server.IpAddress;
            server.Location = virtual_server.Location; server.Status = virtual_server.State;
            server.ActiveDirecotryDomain = virtual_server.ActiveDirectoryDomain;
            if (ping)
            {
                virtual_server.TryPing();
            }
            else
            {
                virtual_server.Ping = "--";
            }
            VirtualizedServer_Result server_ping = new VirtualizedServer_Result();
            server_ping.COMMON_Server = server;
            server_ping.Ping = virtual_server.Ping;
            try
            {
                switch (module)
                {
                    case HomeController.AD_MODULE:
                        if (virtual_server.Site == null || virtual_server.Site == "N/A" || virtual_server.Site.Trim() == "")
                        {
                            virtual_server.SetSite();
                        }
                        FaultyServer ad_server = CastIntoAdServer(server);
                        ad_server.IdSite = (virtual_server.IdSite != null) ? virtual_server.IdSite : "N/A";
                        ad_server.Site = virtual_server.Site;
                        server_ping.AD_Server = ad_server;
                        break;
                    case HomeController.BESR_MODULE:
                        BackupServer besr_server = CastIntoBesrServer(server);
                        besr_server.Version = GetOs(besr_server.Name);
                        server_ping.BESR_Server = besr_server;
                        break;
                    case HomeController.APP_MODULE:
                        AppServer app_server = CastIntoAppServer(server);
                        server_ping.APP_Server = app_server;
                        break;
                    case HomeController.SPACE_MODULE:
                        SpaceServer space_server = CastIntoSpaceServer(server);
                        server_ping.SPACE_Server = space_server;
                        break;
                    default: break;
                }
            }
            catch { }
            return server_ping;
        }

        public static Dictionary<int, VirtualizedServer> GetInformationsFromForestDomains()
        {
            Dictionary<int, VirtualizedServer> domainControllerInfo = new Dictionary<int, VirtualizedServer>();
            Forest obj = System.DirectoryServices.ActiveDirectory.Forest.GetCurrentForest();
            DomainCollection collection = obj.Domains;
            int index = 0;
            try
            {
                foreach (Domain domain in collection)
                {
                    DirectoryContext context = new DirectoryContext(DirectoryContextType.Domain, domain.Name);
                    DomainControllerCollection controllers = System.DirectoryServices.ActiveDirectory.DomainController.FindAll(context);
                    foreach (DomainController controller in controllers)
                    {
                        IPAddress ipAddress = null;
                        VirtualizedServer serverinfo = new VirtualizedServer();
                        string name = controller.Name.Split('.')[0].ToUpper();
                        serverinfo.Name = name;
                        IPAddress.TryParse(controller.IPAddress, out ipAddress);
                        serverinfo.IpAddress = (ipAddress != null) ? ipAddress.ToString() : "N/A";
                        serverinfo.ActiveDirectoryDomain = domain.Name;
                        serverinfo.Site = (controller.SiteName != null && controller.SiteName.Trim() != "") ? controller.SiteName : "N/A"; ;
                        //serverinfo.Version =  controller.OSVersion
                        domainControllerInfo.Add(index, serverinfo);
                        index++;
                    }
                }
            }
            catch { }
            return domainControllerInfo;
        }

        public static FaultyServer CastIntoAdServer(Server server)
        {
            FaultyServer faultyserver = new FaultyServer();
            faultyserver.Id = server.Id;
            faultyserver.Name = server.Name;
            faultyserver.IpAddress = (server.IpAddress != null && server.IpAddress.Trim() != "") ?
                server.IpAddress : "N/A";
            faultyserver.Location = (server.Location != null && server.Location.Trim() != "") ?
                server.Location : "N/A";
            faultyserver.Status = (server.Status != null && server.Status.Trim() != "") ?
                server.Status : "N/A";
            faultyserver.ActiveDirecotryDomain = (server.ActiveDirecotryDomain != null && server.ActiveDirecotryDomain.Trim() != "") ?
                server.ActiveDirecotryDomain : "N/A";
            faultyserver.IdSite = "N/A"; faultyserver.Site = "N/A";
            faultyserver.Version = "N/A";
            return faultyserver;
        }

        public static BackupServer CastIntoBesrServer(Server server)
        {
            BackupServer backupserver = new BackupServer();
            backupserver.Id = server.Id;
            backupserver.Name = server.Name;
            backupserver.IpAddress = (server.IpAddress != null && server.IpAddress.Trim() != "") ?
                server.IpAddress : "N/A";
            backupserver.Location = (server.Location != null && server.Location.Trim() != "") ?
                server.Location : "N/A";
            backupserver.Status = (server.Status != null && server.Status.Trim() != "") ?
                server.Status : "N/A";
            backupserver.ActiveDirecotryDomain = (server.ActiveDirecotryDomain != null && server.ActiveDirecotryDomain.Trim() != "") ?
                server.ActiveDirecotryDomain : "N/A";
            backupserver.Version = (server.Version != null && server.Version.Trim() != "") ?
                server.Version : "N/A";
            backupserver.Disks = "";
            return backupserver;
        }

        public static AppServer CastIntoAppServer(Server server)
        {
            AppServer appserver = new AppServer();
            appserver.Id = server.Id;
            appserver.Name = server.Name;
            appserver.IpAddress = (server.IpAddress != null && server.IpAddress.Trim() != "") ?
                server.IpAddress : "N/A";
            appserver.Location = (server.Location != null && server.Location.Trim() != "") ?
                server.Location : "N/A";
            appserver.Status = (server.Status != null && server.Status.Trim() != "") ?
                server.Status : "N/A";
            appserver.ActiveDirecotryDomain = (server.ActiveDirecotryDomain != null && server.ActiveDirecotryDomain.Trim() != "") ?
                server.ActiveDirecotryDomain : "N/A";
            appserver.Version = (server.Version != null && server.Version.Trim() != "") ?
                server.Version : "N/A";
            return appserver;
        }

        public static SpaceServer CastIntoSpaceServer(Server server)
        {
            SpaceServer spaceserver = new SpaceServer();
            spaceserver.Id = server.Id;
            spaceserver.Name = server.Name;
            spaceserver.IpAddress = (server.IpAddress != null && server.IpAddress.Trim() != "") ?
                server.IpAddress : "N/A";
            spaceserver.Location = (server.Location != null && server.Location.Trim() != "") ?
                server.Location : "N/A";
            spaceserver.Status = (server.Status != null && server.Status.Trim() != "") ?
                server.Status : "N/A";
            spaceserver.ActiveDirecotryDomain = (server.ActiveDirecotryDomain != null && server.ActiveDirecotryDomain.Trim() != "") ?
                server.ActiveDirecotryDomain : "N/A";
            spaceserver.Version = (server.Version != null && server.Version.Trim() != "") ?
                server.Version : "N/A";
            spaceserver.IsShare = false;
            spaceserver.Disks = "";
            return spaceserver;
        }

        [SuppressMessage("Microsoft.Security", "CA2118:ReviewSuppressUnmanagedCodeSecurityUsage"), SuppressUnmanagedCodeSecurity]
        [DllImport("Kernel32", SetLastError = true, CharSet = CharSet.Auto)]
        [return: MarshalAs(UnmanagedType.Bool)]
        private static extern bool GetDiskFreeSpaceEx
        (
            string lpszPath,                    // Must name a folder, must end with '\'.
            ref long lpFreeBytesAvailable,
            ref long lpTotalNumberOfBytes,
            ref long lpTotalNumberOfFreeBytes
        );

        protected override void Dispose(bool disposing)
        {
            db.Dispose();
            base.Dispose(disposing);
        }
    }
}
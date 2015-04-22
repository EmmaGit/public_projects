﻿using System;
using System.Management;
using System.Runtime.InteropServices;
using System.ServiceProcess;

namespace McoEasyTool.Controllers
{
    public class ServiceSpecialController : ServiceController
    {
        public ServiceSpecialController() : base() { }

        public ServiceSpecialController(string name) : base(name) { }

        public ServiceSpecialController(string name, string machineName) : base(name, machineName) { }


        [DllImport("advapi32.dll", SetLastError = true)]
        public static extern bool LogonUser(
                string lpszUsername,
                string lpszDomain,
                string lpszPassword,
                int dwLogonType,
                int dwLogonProvider,
                out IntPtr phToken);

        public enum LogonType
        {
            /// <summary>
            /// This logon type is intended for users who will be interactively using the computer, such as a user being logged on  
            /// by a terminal server, remote shell, or similar process.
            /// This logon type has the additional expense of caching logon information for disconnected operations; 
            /// therefore, it is inappropriate for some client/server applications,
            /// such as a mail server.
            /// </summary>
            LOGON32_LOGON_INTERACTIVE = 2,

            /// <summary>
            /// This logon type is intended for high performance servers to authenticate plaintext passwords.

            /// The LogonUser function does not cache credentials for this logon type.
            /// </summary>
            LOGON32_LOGON_NETWORK = 3,

            /// <summary>
            /// This logon type is intended for batch servers, where processes may be executing on behalf of a user without 
            /// their direct intervention. This type is also for higher performance servers that process many plaintext
            /// authentication attempts at a time, such as mail or Web servers. 
            /// The LogonUser function does not cache credentials for this logon type.
            /// </summary>
            LOGON32_LOGON_BATCH = 4,

            /// <summary>
            /// Indicates a service-type logon. The account provided must have the service privilege enabled. 
            /// </summary>
            LOGON32_LOGON_SERVICE = 5,

            /// <summary>
            /// This logon type is for GINA DLLs that log on users who will be interactively using the computer. 
            /// This logon type can generate a unique audit record that shows when the workstation was unlocked. 
            /// </summary>
            LOGON32_LOGON_UNLOCK = 7,

            /// <summary>
            /// This logon type preserves the name and password in the authentication package, which allows the server to make 
            /// connections to other network servers while impersonating the client. A server can accept plaintext credentials 
            /// from a client, call LogonUser, verify that the user can access the system across the network, and still 
            /// communicate with other servers.
            /// NOTE: Windows NT:  This value is not supported. 
            /// </summary>
            LOGON32_LOGON_NETWORK_CLEARTEXT = 8,

            /// <summary>
            /// This logon type allows the caller to clone its current token and specify new credentials for outbound connections.
            /// The new logon session has the same local identifier but uses different credentials for other network connections. 
            /// NOTE: This logon type is supported only by the LOGON32_PROVIDER_WINNT50 logon provider.
            /// NOTE: Windows NT:  This value is not supported. 
            /// </summary>
            LOGON32_LOGON_NEW_CREDENTIALS = 9,
        }

        public enum LogonProvider
        {
            /// <summary>
            /// Use the standard logon provider for the system. 
            /// The default security provider is negotiate, unless you pass NULL for the domain name and the user name 
            /// is not in UPN format. In this case, the default provider is NTLM. 
            /// NOTE: Windows 2000/NT:   The default security provider is NTLM.
            /// </summary>
            LOGON32_PROVIDER_DEFAULT = 0,
            LOGON32_PROVIDER_WINNT35 = 1,
            LOGON32_PROVIDER_WINNT40 = 2,
            LOGON32_PROVIDER_WINNT50 = 3
        }

        public string Description
        {
            get
            {
                string path = "Win32_Service.Name='" + this.ServiceName + "'";
                ManagementPath p = new ManagementPath(path);

                ManagementObject ManagementObj = new ManagementObject(p);
                if (ManagementObj["Description"] != null)
                {
                    return ManagementObj["Description"].ToString();
                }
                else
                {
                    return null;
                }
            }
        }

        public string StartupType
        {
            get
            {
                if (this.ServiceName != null)
                {
                    var secure = new System.Security.SecureString();
                    foreach (char c in McoUtilities.Decrypt(HomeController.DEFAULT_PASSWORD_IMPERSONNATION))
                    {
                        secure.AppendChar(c);
                    }
                    ConnectionOptions connection = new ConnectionOptions();
                    connection.Username = HomeController.DEFAULT_DOMAIN_AND_USERNAME_IMPERSONNATION;
                    connection.SecurePassword = secure;
                    connection.EnablePrivileges = true;
                    ManagementScope scope = new ManagementScope(String.Format("\\\\{0}\\root\\cimv2", this.MachineName), connection);
                    string path = String.Format("\\\\{0}\\root\\cimv2:Win32_Service.Name='{1}'", this.MachineName, this.ServiceName);
                    ManagementPath p = new ManagementPath(path);
                    ManagementObject ManagementObj = new ManagementObject(scope, p, new ObjectGetOptions(null, System.TimeSpan.MaxValue, true));
                    return ManagementObj["StartMode"].ToString();
                }
                else
                {
                    return null;
                }
            }
            set
            {
                if (value != "Automatic" && value != "Manual" && value != "Disabled")
                    throw new Exception("The valid values are Automatic, Manual or Disabled");
                if (this.ServiceName != null)
                {
                    var secure = new System.Security.SecureString();
                    foreach (char c in McoUtilities.Decrypt(HomeController.DEFAULT_PASSWORD_IMPERSONNATION))
                    {
                        secure.AppendChar(c);
                    }
                    try
                    {
                        ConnectionOptions connection = new ConnectionOptions();
                        connection.Username = HomeController.DEFAULT_DOMAIN_AND_USERNAME_IMPERSONNATION;
                        connection.SecurePassword = secure;
                        connection.EnablePrivileges = true;
                        ManagementScope scope = new ManagementScope(String.Format("\\\\{0}\\root\\cimv2", this.MachineName), connection);
                        string path = String.Format("\\\\{0}\\root\\cimv2:Win32_Service.Name='{1}'", this.MachineName, this.ServiceName);
                        ManagementPath p = new ManagementPath(path);
                        ManagementObject ManagementObj = new ManagementObject(scope, p, new ObjectGetOptions(null, System.TimeSpan.MaxValue, true));
                        object[] parameters = new object[1];
                        parameters[0] = value;
                        ManagementObj.InvokeMethod("ChangeStartMode", parameters);
                    }
                    catch (Exception exception)
                    {
                        McoUtilities.General_Logging(exception, "StartupType ServiceSpecial");
                    }
                }
            }
        }

        public ServiceSpecialController[] GetServices()
        {
            ServiceController[] services = ServiceController.GetServices();
            ServiceSpecialController[] specialServices = new ServiceSpecialController[services.Length];
            int index = 0;
            foreach (ServiceController service in services)
            {
                specialServices[index] = new ServiceSpecialController(service.ServiceName);
            }
            return specialServices;
        }

        public ServiceSpecialController[] GetServices(string machineName)
        {
            ServiceController[] services = ServiceController.GetServices(machineName);
            ServiceSpecialController[] specialServices = new ServiceSpecialController[services.Length];
            int index = 0;
            foreach (ServiceController service in services)
            {
                specialServices[index] = new ServiceSpecialController(service.ServiceName, machineName);
            }
            return specialServices;
        }

    }

}

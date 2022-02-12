using System;
using System.DirectoryServices;
using System.IO;
using System.Security.AccessControl;
using System.Security.Principal;

namespace LDAP.ConsoleApp.NetCore
{
    public class Utils
    {
        public static string? DomainUrl(string? domainUrl) => domainUrl;

        /// <summary>
        /// If you are logged into a system as a domain administrator or 
        /// a user with appropriate privilages then you should 
        /// not need to specify a username and password for the connection.
        /// </summary>
        /// <param name="ldapUrl"></param>
        /// <param name="username"></param>
        /// <param name="password"></param>
        /// <returns></returns>
        public static DirectoryEntry? CreateDirectoryEntry(string? ldapUrl, string? username, string? password) => string.IsNullOrWhiteSpace(username) && !string.IsNullOrWhiteSpace(ldapUrl) ? new DirectoryEntry($"LDAP://{ldapUrl}") : new DirectoryEntry($"LDAP://{ldapUrl}", username, password, AuthenticationTypes.Secure);

        public static int CreateUser(DirectoryEntry? myLdapConnection, string? domain, string? first,
                                    string? last, string? description, object[] password,
                                    string[] groups, string? username, string? homeDrive,
                                    string? homeDir, bool enabled)
        {
            // create new user object and write into AD
            DirectoryEntry? user = myLdapConnection?.Children?.Add($"CN={first} {last}", "user");

            // User name (domain based)   
            user?.Properties["userprincipalname"]?.Add($"{username}@{domain}");

            // User name (older systems)  
            user.Properties["samaccountname"].Add(username);

            // Surname  
            user.Properties["sn"].Add(last);

            // Forename  
            user.Properties["givenname"].Add(first);

            // Display name  
            user.Properties["displayname"].Add($"{first} {last}");

            // Description  
            user.Properties["description"].Add(description);

            // E-mail  
            user.Properties["mail"].Add($"{first}.{last}@{domain}");

            // Home dir (drive letter)  
            user.Properties["homedirectory"].Add(homeDir);

            // Home dir (path)  
            user.Properties["homedrive"].Add(homeDrive);

            user.CommitChanges();

            // set user's password  
            user.Invoke("SetPassword", password);

            // enable account if requested (see http://support.microsoft.com/kb/305144 for other codes)   

            if (enabled)
                user.Invoke("Put", new object[] { "userAccountControl", "512" });

            // add user to specified groups  
            foreach (string? thisGroup in groups)
            {
                DirectoryEntry? newGroup = myLdapConnection.Parent.Children.Find($"CN={thisGroup}", "group");

                if (newGroup != null)
                    newGroup.Invoke("Add", new object[] { user.Path.ToString() });
            }

            user.CommitChanges();

            // make home folder on server  

            Directory.CreateDirectory(homeDir);

            // set permissions on folder, we loop this because if the program  
            // tries to set the permissions straight away an exception will be  
            // thrown as the brand new user does not seem to be available, it takes  
            // a second or so for it to appear and it can then be used in ACLs  
            // and set as the owner  

            bool folderCreated = false;

            while (!folderCreated)
            {
                try
                {
                    // get current ACL  
                    DirectoryInfo? dInfo = new DirectoryInfo(homeDir);
                    DirectorySecurity? dSecurity = dInfo.GetAccessControl();

                    // Add full control for the user and set owner to them  
                    IdentityReference? newUser = new NTAccount($@"{domain}\{username}");

                    dSecurity.SetOwner(newUser);

                    FileSystemAccessRule? permissions = new FileSystemAccessRule(newUser, FileSystemRights.FullControl, AccessControlType.Allow);

                    dSecurity.AddAccessRule(permissions);

                    // Set the new access settings.  
                    dInfo.SetAccessControl(dSecurity);

                    folderCreated = true;
                }
                catch (IdentityNotMappedException)
                {
                    Console.Write(".");
                }
                catch (Exception ex)
                {
                    // other exception caught so not problem with user delay as   
                    // commented above  
                    Console.WriteLine($"Exception caught: {ex}");
                    return 1;
                }
            }

            return 0;
        }
    }
}
using System;
using System.DirectoryServices;
using System.DirectoryServices.AccountManagement;

namespace LDAP.ConsoleApp
{
    internal class Program
    {
        static void Main()
        {
            string ldapUrl, domainUrl, loginUserName = string.Empty, loginPassword = string.Empty;

            Console.Write("--------------------------------------------------------------------------------------------------------------------\n");
            Console.Write("QUERIES ARE SAFE ENOUGH BUT WHEN YOU GET ON TO ACCOUNT CREATION AND MODIFICATION THE POTENTIAL TO ROYALLY MUCK UP A LOT OF ACCOUNT VERY QUICKLY IS A REAL DANGER, SO TAKE CARE!\n");
            Console.Write("--------------------------------------------------------------------------------------------------------------------\n");
            Console.Write("To get started, kindly provide the following:\n\n");

            Console.Write("1. LDAP URL / IP ADDRESS: ");

            ldapUrl = Console.ReadLine();

            if (string.IsNullOrWhiteSpace(ldapUrl))
            {
                throw new ArgumentNullException(nameof(ldapUrl));
            }

            Console.Write("2. Domain URL: ");

            domainUrl = Console.ReadLine();

            if (string.IsNullOrWhiteSpace(domainUrl))
            {
                throw new ArgumentNullException(nameof(domainUrl));
            }

            Console.WriteLine("\nNOTE:If you are logged into a system as a domain administrator or a user with appropriate privilages then you should not need to specify a username and password for the connection.\n");
            Console.Write("3. Login Username (Optional): ");

            loginUserName = Console.ReadLine();

            Console.Write("4. Login Password (Optional): ");

            loginPassword = Console.ReadLine();

            Console.Write($"\nYou entered\nLDAP URL: ('{ldapUrl}'), Domain URL: ('{domainUrl}'), Username: ('{loginUserName}'), Password: ('{loginPassword}')\n\n");

            Console.Write("Select an option to proceed:\n1. All Users\n2. Create User\n3. View Newer Details\n4. View Newer Dropdowns\n5. View Newer Phones\n6. Retrieve All Info\n7. Retrieve Some Info\n8. Update User\n9. View Menu\n00. To enter LDAP Details\n000. To exit\n\n");

            string action;
            do
            {
                Console.Write("Enter Option: ");

                action = Console.ReadLine();

                if (action is null)
                {
                    throw new ArgumentNullException(nameof(action));
                }

                switch (action)
                {
                    case "00":
                        #region Enter URL Details

                        #region Reset Details

                        ldapUrl = string.Empty;
                        domainUrl = string.Empty;
                        loginUserName = string.Empty;
                        loginPassword = string.Empty;

                        #endregion

                        Console.Write("\n1. LDAP URL / IP ADDRESS: ");

                        ldapUrl = Console.ReadLine();

                        if (string.IsNullOrWhiteSpace(ldapUrl))
                        {
                            throw new ArgumentNullException(nameof(ldapUrl));
                        }

                        Console.Write("2. Domain URL: ");

                        domainUrl = Console.ReadLine();

                        if (string.IsNullOrWhiteSpace(domainUrl))
                        {
                            throw new ArgumentNullException(nameof(domainUrl));
                        }

                        Console.WriteLine("\nNOTE:If you are logged into a system as a domain administrator or a user with appropriate privilages then you should not need to specify a username and password for the connection.\n");
                        Console.Write("3. Login Username (Optional): ");

                        loginUserName = Console.ReadLine();

                        Console.Write("4. Login Password (Optional): ");

                        loginPassword = Console.ReadLine();

                        Console.Write($"\nYou entered\nLDAP URL: ('{ldapUrl}'), Domain URL: ('{domainUrl}'), Username: ('{loginUserName}'), Password: ('{loginPassword}')\n\n");

                        #endregion
                        break;
                    case "1":
                        Console.Write("\nAll Users Action\n");
                        AllUsers(ldapUrl, loginUserName, loginPassword);
                        break;
                    case "2":
                        Console.Write("\nCreate User Action\n");
                        CreateUser(ldapUrl, loginUserName, loginPassword, domainUrl);
                        break;
                    case "3":
                        Console.Write("\nNewer Details Action\n");
                        NewerDetails(domainUrl);
                        break;
                    case "4":
                        Console.Write("\nNewer Dropdowns Action\n");
                        NewerDropdowns(domainUrl);
                        break;
                    case "5":
                        Console.Write("\nNewer Phones Action\n");
                        NewerPhones(domainUrl);
                        break;
                    case "6":
                        Console.Write("\nRetrieve All Info Action\n");
                        RetrieveAllInfo(ldapUrl, loginUserName, loginPassword);
                        break;
                    case "7":
                        Console.Write("\nRetrieve Some Info Action\n");
                        RetrieveSomeInfo(ldapUrl, loginUserName, loginPassword);
                        break;
                    case "8":
                        Console.Write("\nUpdate User Action\n");
                        UpdateUser(ldapUrl, loginUserName, loginPassword);
                        break;
                    case "9":
                        Console.Write("Select an option to proceed:\n1. All Users\n2. Create User\n3. View Newer Details\n4. View Newer Dropdowns\n5. View Newer Phones\n6. Retrieve All Info\n7. Retrieve Some Info\n8. Update User\n9. View Menu\n00. To enter LDAP Details\n000. To exit\n\n");
                        break;
                    case "000":
                        break;
                    default:
                        Console.Write("\nInvalid Choice:\n");
                        break;
                }
            } while (action != "000");

            Console.WriteLine("\n\n<~~~~~~~~~~ Good bye Amigos ~~~~~~~~~~~>\n");
            Console.ReadLine();
            Environment.Exit(0);
        }

        #region All Users

        static void AllUsers(string ldapUrl, string loginUserName, string loginPassword)
        {
            Console.Write("Enter property: ");
            string property = Console.ReadLine();

            try
            {
                DirectoryEntry myLdapConnection = Utils.CreateDirectoryEntry(ldapUrl, loginUserName, loginPassword);

                DirectorySearcher search = new DirectorySearcher(myLdapConnection);
                search?.PropertiesToLoad.Add("cn");
                search?.PropertiesToLoad.Add(property);

                SearchResultCollection allUsers = search.FindAll();

                foreach (SearchResult result in allUsers)
                {
                    if (result?.Properties["cn"].Count > 0 && result?.Properties[property].Count > 0)
                    {
                        Console.WriteLine(string.Format("{0,-20} : {1}",
                                      result?.Properties["cn"][0].ToString(),
                                      result?.Properties[property][0].ToString()));
                    }
                }
            }
            catch (Exception e)
            {
                Console.WriteLine($"Exception caught:\n\n{e}\n");
            }
        }

        #endregion

        #region Create User

        static void CreateUser(string ldapUrl, string loginUserName, string loginPassword, string domainUrl)
        {
            // connect to LDAP  

            DirectoryEntry myLdapConnection = Utils.CreateDirectoryEntry(ldapUrl, loginUserName, loginPassword);

            // define vars for user  
            string domain = Utils.DomainUrl(domainUrl);
            string first = "Kelsey";
            string last = "Doe";
            string description = "SPS-SBM Test";
            object[] password = { "12345678" };
            string[] groups = { "Staff" };
            string username = first.ToLower() + last.Substring(0, 1).ToLower();
            Console.Write("Enter the home drive letter: ");
            string homeDrive = $"{Console.ReadLine()}:";
            string homeDir = $@"\\matukadevs.{Utils.DomainUrl(domainUrl)}\data3\USERS\" + username;

            // create user  

            try
            {
                if (Utils.CreateUser(myLdapConnection, domain, first, last, description, password, groups, username, homeDrive, homeDir, true) == 0)
                {
                    Console.WriteLine("Account created!");
                    Console.ReadLine();
                }
                else
                {
                    Console.WriteLine("Problem creating account :(");
                    Console.ReadLine();
                }
            }
            catch (Exception e)
            {
                Console.WriteLine($"Exception caught:\n\n{e}\n");
                Console.ReadLine();
            }
        }

        #endregion

        #region Newer Details

        static void NewerDetails(string domainUrl)
        {
            try
            {
                // enter AD settings  
                PrincipalContext AD = new PrincipalContext(ContextType.Domain, Utils.DomainUrl(domainUrl));

                // create search user and add criteria  
                Console.Write("Enter logon name: ");
                UserPrincipal u = new UserPrincipal(AD)
                {
                    SamAccountName = Console.ReadLine()
                };

                // search for user  
                PrincipalSearcher search = new PrincipalSearcher(u);
                UserPrincipal result = (UserPrincipal)search.FindOne();
                search.Dispose();

                // show some details  
                Console.WriteLine("Display Name : " + result.DisplayName);
                Console.WriteLine("Phone Number : " + result.VoiceTelephoneNumber);
            }
            catch (Exception e)
            {
                Console.WriteLine("Error: " + e.Message);
            }
        }

        #endregion

        #region Newer Dropdowns

        static void NewerDropdowns(string domainUrl)
        {
            try
            {
                PrincipalContext AD = new PrincipalContext(ContextType.Domain, Utils.DomainUrl(domainUrl));
                UserPrincipal u = new UserPrincipal(AD);
                PrincipalSearcher search = new PrincipalSearcher(u);

                foreach (UserPrincipal result in search.FindAll())
                {
                    if (result.VoiceTelephoneNumber != null)
                    {
                        DirectoryEntry lowerLdap = (DirectoryEntry)result.GetUnderlyingObject();

                        Console.WriteLine("{0,30} {1} {2}",
                            result.DisplayName,
                            result.VoiceTelephoneNumber,
                            lowerLdap.Properties["postofficebox"][0].ToString());
                    }
                }

                search.Dispose();
            }
            catch (Exception e)
            {
                Console.WriteLine("Error: " + e.Message);
            }
        }

        #endregion

        #region Newer Phones

        static void NewerPhones(string domainUrl)
        {
            try
            {
                PrincipalContext AD = new PrincipalContext(ContextType.Domain, Utils.DomainUrl(domainUrl));
                UserPrincipal u = new UserPrincipal(AD);
                PrincipalSearcher search = new PrincipalSearcher(u);

                foreach (UserPrincipal result in search.FindAll())
                    if (result.VoiceTelephoneNumber != null)
                        Console.WriteLine("{0,30} {1} ", result.DisplayName, result.VoiceTelephoneNumber);

                search.Dispose();
            }
            catch (Exception e)
            {
                Console.WriteLine("Error: " + e.Message);
            }
        }

        #endregion

        #region Retrieve All Info

        static void RetrieveAllInfo(string ldapUrl, string loginUserName, string loginPassword)
        {
            Console.Write("Enter user: ");
            string username = Console.ReadLine();

            try
            {
                // create LDAP connection object  

                DirectoryEntry myLdapConnection = Utils.CreateDirectoryEntry(ldapUrl, loginUserName, loginPassword);

                // create search object which operates on LDAP connection object  
                // and set search object to only find the user specified  

                DirectorySearcher search = new DirectorySearcher(myLdapConnection)
                {
                    Filter = $"(cn={username})"
                };

                // create results objects from search object  

                SearchResult result = search.FindOne();

                if (result != null)
                {
                    // user exists, cycle through LDAP fields (cn, telephonenumber etc.)  

                    ResultPropertyCollection fields = result?.Properties;

                    foreach (string ldapField in fields.PropertyNames)
                    {
                        // cycle through objects in each field e.g. group membership  
                        // (for many fields there will only be one object such as name)  

                        foreach (object myCollection in fields[ldapField])
                            Console.WriteLine(string.Format("{0,-20} : {1}",
                                          ldapField, myCollection.ToString()));
                    }
                }
                else
                {
                    // user does not exist  
                    Console.WriteLine("User not found!");
                }
            }
            catch (Exception e)
            {
                Console.WriteLine($"Exception caught:\n\n{e}\n");
            }
        }

        #endregion

        #region Retrieve Some Info

        static void RetrieveSomeInfo(string ldapUrl, string loginUserName, string loginPassword)
        {
            Console.Write("Enter user: ");
            string username = Console.ReadLine();

            try
            {
                DirectoryEntry myLdapConnection = Utils.CreateDirectoryEntry(ldapUrl, loginUserName, loginPassword);
                DirectorySearcher search = new DirectorySearcher(myLdapConnection)
                {
                    Filter = $"(cn={username})"
                };

                // create an array of properties that we would like and  
                // add them to the search object  

                string[] requiredProperties = new string[] { "cn", "postofficebox", "mail" };

                foreach (string property in requiredProperties)
                    search?.PropertiesToLoad.Add(property);

                SearchResult result = search.FindOne();

                if (result != null)
                {
                    foreach (string property in requiredProperties)
                        foreach (object myCollection in result?.Properties[property])
                            Console.WriteLine(string.Format("{0,-20} : {1}",
                                          property, myCollection.ToString()));
                }
                else Console.WriteLine("User not found!");
            }
            catch (Exception e)
            {
                Console.WriteLine($"Exception caught:\n\n{e}\n");
            }
        }

        #endregion

        #region Update User

        static void UpdateUser(string ldapUrl, string loginUserName, string loginPassword)
        {
            Console.Write("Enter user: ");
            string username = Console.ReadLine();

            try
            {
                DirectoryEntry myLdapConnection = Utils.CreateDirectoryEntry(ldapUrl, loginUserName, loginPassword);

                DirectorySearcher search = new DirectorySearcher(myLdapConnection)
                {
                    Filter = $"(cn={username})"
                };

                search?.PropertiesToLoad.Add("title");

                SearchResult result = search.FindOne();

                if (result != null)
                {
                    // create new object from search result  

                    DirectoryEntry entryToUpdate = result.GetDirectoryEntry();

                    // show existing title  

                    Console.WriteLine($"Current title   : {entryToUpdate.Properties["title"][0]}");

                    Console.Write("\n\nEnter new title : ");

                    // get new title and write to AD  

                    string newTitle = Console.ReadLine();

                    entryToUpdate.Properties["title"].Value = newTitle;

                    entryToUpdate.CommitChanges();

                    Console.WriteLine("\n\n...new title saved");
                }

                else Console.WriteLine("User not found!");
            }
            catch (Exception e)
            {
                Console.WriteLine($"Exception caught:\n\n{e}\n");
            }
        }

        #endregion

    }
}

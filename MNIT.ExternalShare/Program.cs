using System;
using System.Collections.Generic;
using System.IO;
using System.Security;
using Microsoft.SharePoint.Client;
using Microsoft.SharePoint.Client.Sharing;

namespace MNIT.ExternalShare
{
    class Program
    {
        static void Main(string[] args)
        {
            if (args[0].Length > 1 && args[0].Contains("https"))
            {

                // get credentials from user input to work against O365 site
                ConsoleColor defaultForeground = Console.ForegroundColor;

                // User Enters login name
                Console.ForegroundColor = ConsoleColor.Green;
                Console.WriteLine("Enter your login name");

                Console.ForegroundColor = defaultForeground;
                string userLoginName;
                userLoginName = Console.ReadLine();

                // User Enters password
                Console.ForegroundColor = ConsoleColor.Green;
                Console.WriteLine("Enter your password");

                Console.ForegroundColor = defaultForeground;
                //string userPassword;
                SecureString userPassword = GetPasswordFromConsoleInput();
                // user domain will be empty in an O365 environment
                string userDomain = "";

                // Call the ConsoleSpinner class to let people know that something is happening
                Console.Write("Working...");
                // Build the user object
                ActingUser actingUser = new ActingUser(userLoginName, userPassword, userDomain);

                string siteAddress = args[0];
                string inputFile = "c:\\temp\\ExternalUserList.csv";
                //string[] readUrls = System.IO.File.ReadAllLines("c:\\temp\\ExternalUserList.csv");
                

                //
                // Read in a file line-by-line, and store it all in a List.
                //
                List<string> list = new List<string>();
                try
                {
                    using (StreamReader reader = new StreamReader(inputFile))
                    {
                        string line;
                        while ((line = reader.ReadLine()) != null)
                        {
                            // Add each email address to the virtual list
                            list.Add(line);
                        }
                    }
                }
                catch (Exception ex01Exception)
                {
                    Console.WriteLine(ex01Exception.Message);
                }


                //foreach (string readCurrentLine in readUrls)
                foreach (string readCurrentLine in list)
                {
                    if (!string.IsNullOrEmpty(readCurrentLine.Trim()))
                    {
                        string currentLine = readCurrentLine.Trim();
                        //ShareSite(siteAddress, currentLine, actingUser);
                        SendInvite(siteAddress, currentLine, actingUser);
                        
                        // Add user to a group
                        string groupName = "TestAKK01 Visitors";
                        string action = "add";

                        string[] passAddUserArgs = new string[4];
                        passAddUserArgs[0] = siteAddress;
                        passAddUserArgs[1] = currentLine;
                        passAddUserArgs[2] = groupName;
                        passAddUserArgs[3] = action;
                        AddUserToGroup(passAddUserArgs, actingUser);
                    }
                }
                Console.WriteLine("Send invitation function is complete.");
            }
        }

        public static void SendInvite(string siteAddress, string externalUserEmail, ActingUser actingUser)
        {
            using (var ctx = new ClientContext(siteAddress))
            {
                try
                {
                    ctx.Credentials = new SharePointOnlineCredentials(actingUser.UserLoginName, actingUser.UserPassword);

                    var users = new List<UserRoleAssignment>();
                    users.Add(new UserRoleAssignment()
                    {
                        UserId = externalUserEmail,
                        Role = Role.View
                    });
                    //var messageBody = "This message is for MN.IT customers that use SharePoint sites, help desk staff, SharePoint administrators, CIOs, and MN.IT executive team members. " +
                    //    "The current service is being upgraded and this bulletin outlines the changes you can expect to see, and actions you may need to take.  <br />" +
                    //    "<b>About the service and the upgrade</b><br />" +
                    //    "On (Date TBD), Microsoft will upgrade our current hosted SharePoint sites, to the Office 365 multi-tenant service. " +
                    //    "This upgrade will enhance functionality as described in the details following this message. <br />" +
                    //    "<b>What should I expect during the upgrade window?</b><br />" +
                    //    "* If you are receiving this email, it is because you have been identified as a user of a State of Minnesota SharePoint user. <br />" +
                    //    "* Lorem ipsum dolor sit amet, consectetur adipiscing elit, sed do eiusmod tempor incididunt ut labore et dolore magna aliqua: <br />" +
                    //    "<ul><li>o Ut enim ad minim veniam, quis nostrud exercitation ullamco laboris nisi ut aliquip ex ea commodo consequat.  </li>" +
                    //    "<li>o After they click on the attachment, they will be prompted to sign in with a Microsoft account or an Office 365 organizational account. " +
                    //    "If they do not have a Microsoft Account, the sign-in page has instructions about how to create a Microsoft Account and password. " +
                    //    "(For more information about Microsoft accounts, visit http://windows.microsoft.com/en-US/windows-live/sign-in-what-is-microsoft-account) </li>" +
                    //    "<li>o After the recipient successfully signs in, Duis aute irure dolor in reprehenderit in voluptate velit esse cillum dolore eu fugiat nulla pariatur. </li></ul><br />" +
                    //    "* Excepteur sint occaecat cupidatat non proident, sunt in culpa qui officia deserunt mollit anim id est laborum.<br />";

                    //WebSharingManager.UpdateWebSharingInformation(ctx, ctx.Web, users, true, messageBody, true, true);
                    WebSharingManager.UpdateWebSharingInformation(ctx, ctx.Web, users, true, null, true, true);
                    ctx.ExecuteQuery();
                }
                catch (Exception ex02Exception)
                {
                    Console.WriteLine(ex02Exception);
                }
            }
        }

        private static void ShareSite(string siteAddress, string externalUserEmail, ActingUser actingUser)
        {
            ClientContext ctx = new ClientContext(siteAddress);
            Site siteCollection = ctx.Site;
            ctx.Credentials = new SharePointOnlineCredentials(actingUser.UserLoginName, actingUser.UserPassword);
            ClientRuntimeContext runtimeContext = (ClientContext)siteCollection.Context;
            try
            { 
                var peoplePickerValue = externalUserEmail;
                string roleValue = "group:7"; // int depends on the group IDs at site
                int groupId = 0;
                bool propageAcl = false; // Not relevant for external accounts
                bool sendEmail = true;
                bool includedAnonymousLinkInEmail = false;
                string emailSubject = null;
                string emailBody = "Site shared";
                //UserSharingResult
                SharingResult result = Web.ShareObject(runtimeContext, siteAddress, peoplePickerValue,
                    roleValue, groupId, propageAcl,
                    sendEmail, includedAnonymousLinkInEmail,
                    emailSubject, emailBody);
                ctx.Load(result);
                ctx.ExecuteQuery();
                Console.WriteLine("Email sent to {0}", externalUserEmail);
            }
            catch (Exception ex03Exception)
            {
                Console.WriteLine(ex03Exception);
            }
            finally
            {
                ctx.Dispose();
            }
        }

        private static void AddUserToGroup(string[] args, ActingUser actingUser)
        {
            string siteAddress = args[0];
            string externalUserEmail = args[1];
            string groupName = args[2];
            string action = args[3];
            ClientContext ctx = new ClientContext(siteAddress);
            Site siteCollection = ctx.Site;
            Web currentWeb = ctx.Web;
            ctx.Credentials = new SharePointOnlineCredentials(actingUser.UserLoginName, actingUser.UserPassword);
            try
            {
                
                ctx.Load(currentWeb, w => w.AllProperties, w => w.SiteGroups, w => w.SiteUserInfoList, w => w.Webs,w => w.Title);
                ctx.ExecuteQuery();
                GroupCollection groupCollection = currentWeb.SiteGroups;
                ctx.Load(groupCollection, groups => groups.Include(grps => grps.Users, grps => grps.Title));
                ctx.ExecuteQuery();
                User user;
                var rootWeb = siteCollection.RootWeb;
                var usr = rootWeb.EnsureUser(externalUserEmail);
                ctx.Load(usr);
                ctx.ExecuteQuery();
                Console.WriteLine("User: {0} Login name: {1} Email: {2}",
                        usr.Title, usr.LoginName, usr.Email);
                foreach (Group group in groupCollection)
                {
                   if (group.Title.Equals(groupName))
                   {
                      UserCreationInformation userInfo = new UserCreationInformation();
                      userInfo.LoginName = externalUserEmail;
             
                      if (action == "add")
                      {
                         user = group.Users.Add(userInfo);
                         group.Users.AddUser(user);
                       }
                       else if (action == "remove")
                       {
                         user = group.Users.GetByLoginName(externalUserEmail);
                         group.Users.Remove(user);
                        }
                        group.Update();
                        currentWeb.Update();
                        ctx.ExecuteQuery();
                     }
                  }
                Console.WriteLine("User {0} added to {1}", externalUserEmail, groupName);
            }
            catch (Exception ex04Exception)
            {
                Console.WriteLine(ex04Exception);
            }
            finally
            {
                ctx.Dispose();
            }
        }

        // hiding and recovering secure string for login
        public static SecureString GetPasswordFromConsoleInput()
        {
            ConsoleKeyInfo info;
            //Get the user's password as a SecureString
            SecureString securePassword = new SecureString();
            do
            {
                info = Console.ReadKey(true);
                if (info.Key != ConsoleKey.Enter)
                {
                    securePassword.AppendChar(info.KeyChar);
                }
            }
            while (info.Key != ConsoleKey.Enter);
            return securePassword;
        }
    }

    public class ActingUser
    {
        public string UserLoginName { get; private set; }
        public SecureString UserPassword { get; private set; }
        public string UserDomain { get; private set; }

        public ActingUser(string userLoginName, SecureString userPassword, string userDomain)
        {
            UserLoginName = userLoginName;
            UserPassword = userPassword;
            UserDomain = userDomain;
        }
    }
}

# InviteExternalUsers
Console app to send invite email to external users to share SharePoint site
You need to have a file called ExternalUserList.csv in the C:\temp directory with one email address per line and nothing else - not error checking at this point yet.
Run from command line. An example of use would be like:
CD C:\Temp\ExternalShare
C:
MNIT.ExternalShare.exe "https://mtSiteName.sharepoint.com/sites/testSite"
Then you will be prompted for username and password to authenticate against the MT O365 site.
Current issues yet to be resolved:
1. I can't seem to figure out how to change the body of the email into rich text.
2. I need to understand how to change who the email is being sent from - I don't want replies in my personal inbox from 2000 external recipients with their questions.
3. I can either send an email using WebSharingManager.UpdateWebSharingInformation, or it seems like I can build a JSON object using Vesa's example (https://blogs.msdn.microsoft.com/vesku/2015/10/02/external-sharing-api-for-sharepoint-and-onedrive-for-business/), but that doesn't seem to send an invite.  Am I missing something?

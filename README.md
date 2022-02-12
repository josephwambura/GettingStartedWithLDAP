# Active Directory With C#

If you have a client in the kind of large institution that I do and they are using Microsoft Active Directory then the chances are that at certain times you will need to perform actions on the directory that are outside the scope of the MSAD tools. This could be things like specialised queries, bulk account creation or mass updates of user information. The MSAD tools and even some of the command line tools are quite limiting and difficult to use in this regard.

Whatever the reason, you may find that at some point you need to either purchase additional software for managing AD or write your own. Obviously I’d rather write my own software as it’s cheaper, more rewarding and you can customise it however you like!

I have combined the following these two ways to achieve this:

* System.DirectoryServices

http://msdn.microsoft.com/en-us/library/system.directoryservices.aspx

The properties of the AD objects (description, telephone etc.) are all held in an array which can present its own problems and involve a lot of iteration and use of casting since they are all generic objects.

Using this approach doing things like setting the password or enabling/disbaling the account is much more cryptic in the way in which it is achieved, often requiring UAC codes to be manually set and so on.

* System.DirectoryServices.AccountManagement

http://msdn.microsoft.com/en-us/library/system.directoryservices.accountmanagement.aspx

Here, Rather than accessing properties using an array they are exposed directly within the classes (and typed accordingly), allowing us to use things like user.DisplayName which is much tidier.

We also have easy to use methods available such as .SetPassword() and .UnlockAccount() as well as the .Enabled property which can be used to easily manage accounts. These are self explanatory in use once you have retrieved the object from AD so are not included in the examples!

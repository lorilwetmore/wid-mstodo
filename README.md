# mstodo
Widget for [smashing](https://github.com/Smashing/smashing) dashboard, accessing Microsoft To Do

## Description


## Dependencies

Add the following to your Gemfile

    gem 'rest-client'
    gem 'json'

and run 'bundle-install'

You will need an application registered at [Microsoft Azure AD](https://portal.azure.com/#blade/Microsoft_AAD_IAM/ActiveDirectoryMenuBlade/RegisteredApps).
I have created an "Application from personal account" which provides a client ID and client secret that you will both need later.


## Setup

Put the files in the respective folders of your smashing project.

Create a new file "assets/config/msauth_settings.json" with the following content

    {"clientsecret":"<YOUR CLIENT SECRET>","clientid":"<YOUR CLIENT ID>"}

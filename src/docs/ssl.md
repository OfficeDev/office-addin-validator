# Adding Self-Signed Certificates as Trusted Root Certificate

Office clients require add-ins and webpages to come from a trusted and secure location. This generator leverages [Browsersync](https://browsersync.io/) to start a web server, which requires a self-signed certificate. Your workstation will not trust this certificate and thus, the Office client, in which you are running your Office Add-in, will not load your add-in.

To fix this, you need to configure your developer workstation to trust the self-signed certificate. The steps for this differ depending on your developer environment (OSX / Windows / Linux). Use these instructions to trust the certificate:

- OS X: [Apple Support - OS X Yosemite: If your Certificate Isn't Being Accepted](https://support.apple.com/kb/PH18677)
- Windows: [TechNet - Manage Trusted Root Certificates](https://technet.microsoft.com/en-us/library/cc754841.aspx)

## Trusting a Self-Signed Certificate on OS X using the Chrome browser

Using Chrome, when you browse to a site that has an untrusted certificate, the browser will display an error with the certificate:

  ![](assets/ssl-error.png)

### Option #1: Simply Proceed

choose the “Advanced” link, then choose “Proceed to local (unsafe)“.

  ![](assets/ssl-chrome.gif)

### Option #2: Trusting a certificate

1. Start Chrome and do the following:

  1. Open Developer Tools window by using keybaord shortcuts: Cmd + Opt + I.
  1. Click to go to 'security' panel and 'overview' screen.
	1. Click 'View certificate'.

  ![](assets/ssl-devtool.png)

1. Click and drag the image to your desktop. It looks like a little certificate.

  ![](assets/ssl-get-cert.png)

1. Open the **Keychain Access** utility in OS X.
  1. Select the **System** option on the left.
  1. Click the lock icon in the upper-left corner to enable changes.

    ![](assets/ssl-keychain-01.png)

  1. Click the plus button at the bottom and select the **localhost.cert** file you copied to the desktop.
  1. In the dialog that comes up, click **Always Trust**.
  1. After **localhost** gets added to the **System** keychain, double-click it to open it again.
  1. Expand the **Trust** section and for the first option, pick **Always Trust**.

    ![](assets/ssl-keychain-02.png)

At this point everything has been configured. Quit Chrome and all other browsers and try again to navigate to the local HTTPS site. The browser should report it as a valid certificate:

![](assets/ssl-good.png)

You now have a self-signed certificate installed on your machine.

Copyright (c) 2017 Microsoft Corporation. All rights reserved.

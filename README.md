# SPFx Development Guide: From Microsoft 365 Developer Subscription to SharePoint Online Environment Setup

1. **Obtain a Microsoft 365 Developer Subscription**

- Get a Microsoft 365 Developer Subscription

- [https://developer.microsoft.com/en-us/microsoft-365/dev-program](https://developer.microsoft.com/en-us/microsoft-365/dev-program)

- Create Sample data packs (User, Mail & Event och Sharepoint)

1. **Join the Microsoft 365 Developer Program**

- Set up a Microsoft 365 developer.

- [https://learn.microsoft.com/en-us/office/developer-program/microsoft-365-developer-program](https://learn.microsoft.com/en-us/office/developer-program/microsoft-365-developer-program)

- Set up a developer subscription.

- https://learn.microsoft.com/en-us/office/developer-program/microsoft-365-developer-program-get-started

- set-up-your-development-environment.

- [https://learn.microsoft.com/en-us/sharepoint/dev/spfx/set-up-your-development-environment](https://learn.microsoft.com/en-us/sharepoint/dev/spfx/set-up-your-development-environment)

            1. Install Node.js
            2. Install a code editor.
            3. Install development toolchain prerequisites:

You need to install a few dependencies globally on your workstation npm install gulp-cli yo @microsoft/generator-sharepoint –global

            1. Install Gulp

Gulp is a JavaScript-based task runner used to automate repetitive tasks.

   ```bash
     npm install gulp-cli –global
```

            1. Install Yeoman

Yeoman helps you kick-start new projects and prescribes best practices and tools to help you stay productive.

   ```bash
     npm install yo –global
```

            1. Install Yeoman SharePoint generator.

The Yeoman SharePoint web part generator helps you quickly create a SharePoint client-side solution project.

   ```bash
     npm install @microsoft/generator-sharepoint –global
```

            1. Trusting the self-signed developer certificate

Self-Signed SSL certificates are not trusted by your developer environment. You must first configure your development environment to trust the certificate.c

   ```bash
     gulp trust-dev-cert
```

1. **Build a template web-part.**

### I note that the web part is based on "react": "17.0.1". It uses React and the code has been updated to a functional component. The links referenced in the document are based on a different version of the framework and component.

- Build SharePoint client-side web part.

- [https://learn.microsoft.com/en-us/sharepoint/dev/spfx/web-parts/get-started/build-a-hello-world-web-part](https://learn.microsoft.com/en-us/sharepoint/dev/spfx/web-parts/get-started/build-a-hello-world-web-part)

            1. Create a new project by running the Yeoman SharePoint Generator from within the new directory you created.

    ```bash
     yo @microsoft/sharepoint
   ```

- Which type of client-side component to create? **WebPart**
- What is your Web part name? **Template**
- Which template would you like to use? **React**

            1. open the file ./config/serve.json in your project and update your project's hosted workbench URL.

- build and preview your web part.

    ```bash
     gulp serve
   ```
use gulp as the task runner to handle build process tasks such as:

- Transpile TypeScript files to JavaScript.

- Compile SASS files to CSS.

- Bundle and minify JavaScript and CSS files.

1. **Deploy your client-side web part to a SharePoint page**

- Package the HelloWorld web part

Enter the following command to bundle your client-side solution.

   ```bash
     gulp bundle
   ```

enter the following command to package your client-side solution that contains the web part.

   ```bash
     gulp package-solution
   ```

- The command creates the following package:  **./sharepoint/solution/helloworld-webpart.sppkg.**
- You can view the raw package contents in the **. /sharepoint/solution/debug folder.**

1. **Do not have an app catalog?**

- [https://learn.microsoft.com/en-us/sharepoint/use-app-catalog?redirectSourcePath=%252farticle%252fuse-the-app-catalog-to-make-custom-business-apps-available-for-your-sharepoint-online-environment-0b6ab336-8b83-423f-a06b-bcc52861cba0](https://learn.microsoft.com/en-us/sharepoint/use-app-catalog?redirectSourcePath=%252farticle%252fuse-the-app-catalog-to-make-custom-business-apps-available-for-your-sharepoint-online-environment-0b6ab336-8b83-423f-a06b-bcc52861cba0)

            1. Go to the More features page in the SharePoint admin center and select Open under Apps.
            2. Open Apps.
            3. Back to Apps.
            4. If you see classic experience in the app catalog, choose to move the new experience by clicking Try the new Manage Apps page in the header.
            5. Upload or drag and drop the spfx-webpart.sppkg to the app catalog.

1. **Preview the web part on a SharePoint page**

- Run the gulp task to start serving from localhost.

   ```bash
     gulp serve –nobrowser
   ```

            1. In your browser, go to your site where the solution was installed.
            2. Select the gears icon in the top nav bar on the right, and then select Add a page.
            3. Edit the page.
            4. Open the web part picker and select your HelloWorld web part.

1. **Hostin your client-side web part from Microsoft 365 CDN**

_Important_

_This article uses the includeClientSideAssets attribute, was introduced in the SharePoint Framework (SPFx) v1.4. This version is not supported with SharePoint 2016 Feature Pack 2. If you're using an on-premises setup, you need to decide the CDN hosting location separately. You can also simply host the JavaScript files from a centralized library in your on-premises SharePoint to which your users have access. Please see additional considerations in the SharePoint 2016 specific guidance._

1. **Enable CDN in your Microsoft 365 on your SharePoint Online tenant in Microsoft 365.**

- Options for managing the Microsoft 365 CDN

            1. Microsoft SharePoint Online Management Shell

- This is a PowerShell module specifically focused on managing SharePoint Online.
- It is a powerful tool if you are familiar with PowerShell and prefer working with it.
- It allows you to use SharePoint Online-specific cmdlets to perform tasks unique to SharePoint Online.
- It allows you to use SharePoint Online-specific cmdlets to perform tasks unique to SharePoint Online.

            1. CLI for Microsoft 365:

- This is a cross-platform command-line interface that can be used on Windows, macOS, and Linux.
- It is more general and useful for administering the entire Microsoft 365 environment, not just SharePoint Online.
- It is great if you want to use a single tool to manage multiple parts of your Microsoft 365 environment, including SharePoint Online, Exchange, Teams, and more.
- If you don't have much experience with PowerShell or prefer to work with a cross-platform solution, this can be a good choice.

**Summary** if you primarily focus on managing SharePoint Online and are comfortable with PowerShell, Microsoft SharePoint Online Management Shell may be more suitable. On the other hand, if you need to manage various aspects of Microsoft 365 and want to use the same tool across different platforms, CLI for Microsoft 365 could be a good option. The choice depends on your specific needs and skills.

**Management Shell**

To get started using PowerShell to manage SharePoint Online, you need to install the SharePoint Online Management Shell and connect to SharePoint Online.

- [https://learn.microsoft.com/en-us/powershell/sharepoint/sharepoint-online/connect-sharepoint-online?view=sharepoint-ps&redirectedfrom=MSDN](https://learn.microsoft.com/en-us/powershell/sharepoint/sharepoint-online/connect-sharepoint-online?view=sharepoint-ps&redirectedfrom=MSDN)

- Install the SharePoint Online Management

- [https://www.microsoft.com/en-us/download/details.aspx?id=35588](https://www.microsoft.com/en-us/download/details.aspx?id=35588)

- To connect with a username and password

Connect-SPOService -Url https://contoso-admin.sharepoint.com -Credential [admin@contoso.com](mailto:admin@contoso.com)

- To connect with multifactor authentication (MFA)

Connect-SPOService -Url [https://contoso-admin.sharepoint.com](https://contoso-admin.sharepoint.com/)

You are now ready to use SharePoint Online commands. Set-SPOTenantCdnEnabled -CdnType Public

            1. Enable public CDN in the tenant.
   ```bash
     Set-SPOTenantCdnEnabled -CdnType Public
   ```

            1. Add a new CDN origin In this case, we're setting the origin as \*/cdn, which means that any relative folder with the name of cdn acts as a CDN origin.


   ```bash
     Add-SPOTenantCdnOrigin -CdnType Public -OriginUrl \*/cdn
   ```

            1. Get the Microsoft 365 CDN status.

   ```bash
   Get-SPOTenantCdnEnabled -CdnType Public

Get-SPOTenantCdnOrigins -CdnType Public

Get-SPOTenantCdnPolicies -CdnType Public
   ```

**Note** If you encounter issues trying to connect or receive an error such as 'Error message: Could not connect to SharePoint Online', you may need to use Modern Authentication. See the following example:

Connect-SPOService -Credential $creds -Url https://tenant-admin.sharepoint.com -ModernAuth $true -AuthenticationUrl [https://login.microsoftonline.com/organizationsv](https://login.microsoftonline.com/organizationsv)

            1. Verify the origin was added

   ```bash
     Get-SPOTenantCdnOrigins -CdnType Public
   ```

1. **Prepare web part assets to deploy**

            1. Execute the following task to bundle your solution. This executes a release build of your project by using a dynamic label as the host URL for your assets. This URL is automatically updated based on your tenant CDN settings.

   ```bash
     gulp bundle –ship
   ```

            1. Execute the following task to package your solution. This creates an updated **helloworld-webpart.sppkg** package on the **sharepoint/solution** folder.

   ```bash
     gulp package-solution –ship
   ```

            1. Upload or drag and drop the newly created client-side solution package to the app catalog in your tenant. Select Replace

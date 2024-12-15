# Extortion Pack

## If your M365 governance is "delegated permissions only", you are going to get hacked and it will be ALL YOUR FAULT

If your M365 governance relies solely on delegated permissions, you are exposing your organization to significant security risks. Malicious actors can exploit access tokens and permissions to steal data, replace links with spoofing URLs, and potentially take over the entire tenant.

## Best case? Your data gets stolen. Worst case? You lose your entire tenant.

Let me explain.

SPFx solutions, and any code running within you SPO all have access to all the SPO sites and resources on current user behalf. They also can access any external APIs. No additional permissions needed, no way to block it.

And if the app requests permissions? "All permissions are granted to the whole tenant and not to a specific application that has requested them." See [Connect to Azure AD-secured APIs in SharePoint Framework solutions](https://learn.microsoft.com/en-us/sharepoint/dev/spfx/use-aadhttpclient#:~:text=All%20permissions%20are%20granted%20to%20the%20whole%20tenant%20and%20not%20to%20a%20specific%20application%20that%20has%20requested%20them).

It means that any permissions requested by apps, once approved, are granted to the "**SharePoint Online Client Extensibility Web Application Principal**" (SPOCEWAP).

They are **not granted to the app that requested them**. **All the code** running in your SharePoint online site, be it SharePoint Framework solutions, JavaScript code in script editor, or JavaScript code you execute in browser’s console **has access** to all SharePoint online data the current user has access to, and to any 3rd party API.

This means that the code acting on behalf of current user (you, your colleagues, your CEO) may either steal your company's data or replace some links with spoofing URLs.

## But wait, it gets worse

The code is using [Access Tokens](https://pnp.github.io/blog/post/introduction-to-tokens/) when accessing information. They are saved in your browser's local storage, and you can see them using the DevTools Application tab (F12 to open).

The code running within your SharePoint Online site (SharePoint framework, JavaScript in script editor) may use these access tokens, or generate fresh ones valid for one hour They may be used outside of SharePoint, or even sent outside privileged your tenant.

Bad actor only needs to:

-   read your current access token to understand what privileged roles you have and what actions may be executed on your behalf
-   generate new Access Token
-   send it to their own endpoint
-   use it within one hour

Depending on what delegated permissions you granted to the SPOCEWAP, and what privileged role your signed-in user has, they may now

-   spend the next hour stealing your data and replacing links within your SPO sites
-   create new app registration and grant App-only permissions to be used by their own code from now on, or
-   take over your tenant

## Delegated permissions

I'm not saying there is something wrong with delegated permissions. [App-only and delegated](https://learn.microsoft.com/en-us/entra/identity-platform/permissions-consent-overview) permissions are different tools for different jobs and it is your responsibility to understand and use them appropriately.

But if you assume that granting delegated permissions only will keep your tenant secure you're up to rude awakening.

## If I was a hacker

If I was a hacker I would build a solution consisting of:

-   a Web Part that engages users in some way a table of contents for a page a dashboard
-   an app customizer that does all the naughty things

Maybe you won't deploy my solution to the whole tenant. No problem.

Once added to a page, the Web Part would associate the App Customizer with all the sites you are an Owner of. The App Customizer would then do the same for any other user accessing infected sites. It would spread like a virus.

Soon all the SharePoint online sites in your environment will be infected.

## Would you let a stranger into your own home, just to see what happens?

It's practically impossible to make a code reviews of a SPFx solutions.

You may consider checking the traffic (F12) generated by the solutions you install. Maybe it will help you maybe it won't.

If I was trying to hack you, I would first check if this is your productive tenant (domain name, number of users) or simply grant a grace period to allow you to test and approve the solution before executing the malicious code. It's a "whack-a-mole" game.

And if I saw you are a Global Admin and the `User.ReadWrite.All` plus`RoleManagement.ReadWrite.Directory` delegated permissions have been granted to the SPOCEWAP, I wouldn’t waste my time. I would send your Access Token to my own endpoint and use it immediately to take over your tenant.

**If you see your access token being stolen it's already too late.**

## Don't underestimate the severity of the problem.

Your red lines should be:

-   granting API permissions that allow creating service principles and registrations or users and
-   accessing any SharePoint online site using roles with elevated privileges

Accept that delegated only approach will not protect you. Only install solutions that you really need and that come from legitimate source, it means nothing, that a website looks good. It's a piece of cake to generate a perfect copy of existing site. Do you know the company is it really their own website?

And if you allow site-level app catalog in your company, make sure the Site Owners do understand the risks, and consent to regular app reviews. Productivity trumps security. Don’t expect your users to care about security if you don't.

## Still don’t believe me?

One ~~image~~ WebPart is worth a thousand words. I wrote two, each of them presenting another vector of attack.

This solution does not request any API permissions, instead relying on those already granted. This is a type of an app you would install no questions asked, right?

It's possible that not all actions may be executed. Perhaps the "Data exfiltration" WebPart won't be able to read your emails? Or maybe the Get Access Token will be useless because you are very conservative when granting API permissions? I sure hope so.

You may still install it in your development tenant, grant extra API permissions to the SPOCEWAP and see what happens.

**These WebParts are not malicious**. They will not make any changes to your tenant, exfiltrate your data or steal access tokens.

### Data Exfiltration

The "Data Exfiltration" WebPart calls https://httpbin.org/post endpoint to ensure that calling external endpoints is not blocked. You may also configure this Web Part and provide the URL and subscription key of your API Management service in Azure.
This will enable additional buttons in the SharePoint REST API tab, allowing you to send the contents of a selected site to that endpoint.

### Get Access Tokens

The "Get Access Tokens" WebPart does not allow sending information to an external APIs. It displays new access token generated for different scopes (https://graph.microsoft.com and https://management.azure.com), along with instructions on how to execute attacks.
You can download both WebParts from my GitHub repo in the Releases section.

And… I hope this proposal gave you a pause. Is it safe? Do you even know me? Can you trust me?

The source code is available, you may review it and build your own package by following these instructions: Prepare web part assets to deploy.

**Choose your own adventure.**

## Don’t kill the messenger

SharePoint and SharePoint extensions have been around for so long that, unlike newer technologies like Copilot extensions, they are rarely perceived as potential security threats.

Typically, most solution architects and developers understand the full extent of what these extensions can do in terms of potential damage, but they are generally focused on building safe and reliable solutions, not exploiting them.

On the other hand, the teams responsible for governance such as Global Administrators or SharePoint Administrators, often lack deep understanding of the security model behind SPFx extensions.

They may be under impression that SPFx solutions are "secure by design" and offer them full control over what is happening in their tenant.

This couldn’t be further from the truth. SPFx may steal your data, replace existing links with urls to spoofing websites, and under right circumstances even steal your tenant.

The threat seems to be **hiding in plain sight**.

SharePoint Product Group is fully aware of those risks. They mention them in the official documentation, but they also make a lot of effort not to sound too alarming. They do such a great job that, in fact, hardly anyone realizes the security risks associated with this model.

I know that Andrew Conell raised these issues with the SharePoint product group but nothing changed. Naively, I contacted the Microsoft Security Response Center group and received the following answer:

“Thank you again for submitting this issue to Microsoft. Currently, MSRC prioritizes vulnerabilities that are assessed as Important or Critical severities for immediate servicing. After careful investigation, this case has been assessed and as presented this looks to be part of the current design does not meet MSRCs bar for immediate servicing. The governance is in the hands of administrators that must approve the permissions before any custom code can use it. Here is some documentation on this topic https://learn.microsoft.com/en-us/sharepoint/dev/spfx/use-aadhttpclient . We have shared the report with the team responsible for maintaining the product or service. They will take appropriate action as needed to help keep customers protected.”

## Popular misconceptions

### We grant delegated permissions only

The governance is often based on (and limited to) granting SPFx solutions delegated permissions only. As great as it sounds, it proves lack of understanding of security model, because SPFx solutions may only ever request delegated permissions.

And although these permissions make sure the user may only see resources they have access to, it won't stop malicious code from stealing data from your company, because **there's no way to stop SPFx apps from calling external API** .

### Until you approve the required permissions the solution won’t work the way it’s supposed to.

There's a common (and false) belief that permissions granted to a SPFx solution apply to this specific solution only, and until the request is approved, the app will not function properly.

In fact, API permissions are shared across all SPFx solutions that ever have, and will be, installed.

### It's not so easy to steal a token

Well... from within the SPFx solution? It is very easy indeed.

It's two lines of code using MSAL library, and the token is valid 1 hour (90 minutes, to be exact). It's more than enough time to do some damage.

## The false sense of security

This false sense of security and a belief that SPFx apps are secure by design may result in lack of scrutiny of SPFx extensions installed in the tenant.

I don't think Administrators cannot be blamed here.

The UI of the "API access" site shows which extension requested which permissions, leading admins to believe they are granting permissions to this specific extension. I alredy heard that "we only need to find out which extension is granted which permissions".

The offical documentation is also not very forthcoming.
The [Manage apps using the Apps site](https://learn.microsoft.com/en-us/sharepoint/use-app-catalog) linked from the "Manage apps" does not mention approving permissions at all. The [Manage access to Microsoft Entra ID-secured APIs](https://learn.microsoft.com/en-us/sharepoint/api-access) does, however, explain that:

> Permissions of type delegated are added to the SharePoint Online Client Extensibility Web Application Principal in Microsoft Entra ID.
>
> If you try to approve a permission request for a resource that already has some permissions granted (for example, granting additional permissions to the Microsoft Graph), the requested scopes are added to the previously granted permissions.

**Finally! It is now clear that all SPFx extensions, along with any JavaScript client-side injected into the page or executed from the console share the same permissions. **

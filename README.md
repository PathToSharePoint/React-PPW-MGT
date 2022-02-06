# React-PPW-MGT

## Applies to

- [SharePoint Framework](https://aka.ms/spfx)
- [Microsoft 365 tenant](https://docs.microsoft.com/en-us/sharepoint/dev/spfx/set-up-your-developer-tenant)

> Get your own free development tenant by subscribing to [Microsoft 365 developer program](http://aka.ms/o365devprogram)

## Summary

The React-PPW-MGT sample showcases the use of the [Property Pane Wrap](https://www.npmjs.com/package/property-pane-wrap) to embed [Microsoft Graph Toolkit](https://www.npmjs.com/package/property-pane-wrap) controls in the SPFx Property Pane.

![React PPW MGT Sample](./assets/React-PPW-MGT-Sample.png)
## Compatibility

![Compatible with SharePoint Online](https://img.shields.io/badge/SharePoint%20Online-Compatible-green.svg)
![Does not work with SharePoint 2019](https://img.shields.io/badge/SharePoint%20Server%202019-Incompatible-red.svg "SharePoint Server 2019 requires SPFx 1.4.1 or lower")
![Does not work with SharePoint 2016 (Feature Pack 2)](https://img.shields.io/badge/SharePoint%20Server%202016%20(Feature%20Pack%202)-Incompatible-red.svg "SharePoint Server 2016 Feature Pack 2 requires SPFx 1.1")
![Hosted Workbench Compatible](https://img.shields.io/badge/Hosted%20Workbench-Compatible-green.svg)

## Used SharePoint Framework Version

![1.14.0.beta.5](https://img.shields.io/badge/version-1.14.0.beta.5-green.svg)

## Solution

Solution|Author(s)
--------|---------
React-PPW-MGT | [Christophe Humbert](https://github.com/PathToSharePoint)

## Version history

Version|Date|Comments
-------|----|--------
0.1.0|February 5, 2022|Initial release

## Disclaimer

**THIS CODE IS PROVIDED *AS IS* WITHOUT WARRANTY OF ANY KIND, EITHER EXPRESS OR IMPLIED, INCLUDING ANY IMPLIED WARRANTIES OF FITNESS FOR A PARTICULAR PURPOSE, MERCHANTABILITY, OR NON-INFRINGEMENT.**

---

## Prerequisites

This solution requires the following Microsoft Graph consent, [granted by the SharePoint Online administrator](https://docs.microsoft.com/en-us/sharepoint/dev/spfx/web-parts/get-started/using-microsoft-graph-apis#approve-the-requested-microsoft-graph-permissions):
- User.ReadBasic.All
- People.Read.All
- Group.Read.All

If API permissions are not granted, the controls will be displayed but won't fetch the content.

## Minimal Path to Awesome

- Clone this repository
- Ensure that you are at the solution folder
- in the command-line run:
  - **npm install**
  - **gulp serve**

## Features

The sample web part illustrates the following concepts on top of the SharePoint Framework:
- MGT [SharePointProvider](https://docs.microsoft.com/en-us/graph/toolkit/get-started/build-a-sharepoint-web-part#add-the-sharepoint-provider) for authentication
- [Property Pane Wrap](https://www.npmjs.com/package/property-pane-wrap) for insertion of components in the Property Pane
- [MGT controls](https://mgt.dev/) (React integration) in the Property Pane
- Cascading selection (Group members) in the Property Pane
- MGT themes (light/dark)

## References

- [Getting started with SharePoint Framework](https://docs.microsoft.com/en-us/sharepoint/dev/spfx/set-up-your-developer-tenant)
- [Building for Microsoft teams](https://docs.microsoft.com/en-us/sharepoint/dev/spfx/build-for-teams-overview)
- [Use Microsoft Graph in your solution](https://docs.microsoft.com/en-us/sharepoint/dev/spfx/web-parts/get-started/using-microsoft-graph-apis)
- [Publish SharePoint Framework applications to the Marketplace](https://docs.microsoft.com/en-us/sharepoint/dev/spfx/publish-to-marketplace-overview)
- [Microsoft 365 Patterns and Practices](https://aka.ms/m365pnp) - Guidance, tooling, samples and open-source controls for your Microsoft 365 development
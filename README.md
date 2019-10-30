# Change-OD4B-TimeZone

Dear SharePoint Online administrators,

The default setting of time zone in OneDrive for Business is 13 [(GMT-08:00) Pacific Time (US and Canada)].
If you change your time zone, you should change this setting manually by this [article](https://support.microsoft.com/ja-jp/help/2901182).
So, this script will help you to change all users' time zone in OneDrive for Business in bulk.

## Prerequirements
1. You need to download and install SharePoint Online Management Shell to run this script. https://www.microsoft.com/en-us/download/details.aspx?id=35588

1. You can also acquire the latest SharePoint Online Client SDK by Nuget as well.
   - You need to access the following site. https://www.nuget.org/packages/Microsoft.SharePointOnline.CSOM
   - Download the nupkg.
   - Change the file extension to *.zip.
   - Unzip and extract those file.
   - place "lib" folder under `C:\csom`

## Usage
`.\Change-OD4B-TimeZone -SPOAdminUrl https://contoso-admin.sharepoint.com -SPOAdminUser spoadmin@contoso.onmicrosoft.com [-TimeZoneId 20]`

## Parameter

- `-SPOAdminUrl` Specifiy the URL of tenant's SharePoint Admin Center.
- `-SPOAdminUser` Specify the user who has SharePoint Administrator role.
- `-TimeZoneId` (Optional) Specify the target time zone id. The default value is 20 [(GMT+09:00) Osaka, Sapporo, Tokyo]

See also :
[RegionalSettings.TimeZones property](https://docs.microsoft.com/en-us/previous-versions/office/sharepoint-csom/jj171282(v%3Doffice.15))

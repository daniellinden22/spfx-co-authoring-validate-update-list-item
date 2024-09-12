# cavuli

## Summary
A small solution to explore how to make validateUpdateListItem requests work on SharePoint pages where co-authoring is enabled. 

## Background
On SharePoint pages where co-authoring is enabled, requests to the validateUpdateListItem endpoint will fail with the error message ``The file "https://{SUB_DOMAIN}.sharepoint.com/sites/{SITE}/SitePages/{PAGE_NAME}.aspx" is locked for shared use by {CURRENT_USER}@{SUB_DOMAIN}.onmicrosoft.com.`` and the status code ``423 Locked``.

## Findings
- When editing properties in the Page Details panel, the validateUpdateListItem request is slightly different.
  - The request body contains the ``sharedLockId`` and ``checkInComment`` properties.
- The Page Details panel uses ``https://{SUB_DOMAIN}.sharepoint.com/sites/{SITE}/_api/web/GetList(@a1)/items(@a2)/ValidateUpdateListItem()?@a1='/sites/{SITE}/SitePages'&@a2='{LIST_ITEM_ID}'``, while PnPJs uses ``https://{SUB_DOMAIN}.sharepoint.com/sites/{SITE}/_api/web/lists('{LIST_ID}')/items({LIST_ITEM_ID})/validateupdatelistitem``. 
  - As far as I can tell, the endpoints are the same (except the list part).
- When editing a page, there is a ``spClientSidePageContext`` JSON object in the DOM.
  - This object contains an object called CoAuthState, which in turn contains the ``sharedLockId`` property.
- The ``sharedLockId`` can be retrieved programmatically, via the Site Pages module/component. I haven't found any other way to retrieve it yet.
  - I believe this should be considered unsupported, as I can't find much (any) documentation about it.
- When providing the ``sharedLockId`` in the request body, validateUpdateListItem requests will succeed.
  - This is true for both of the above endpoint, however the PnPJs parameters for their validateUpdateListItem method does not include the ``sharedLockId`` property, nor the ``checkInComment``.
  - My understanding is that the PnPJs implementation of validateUpdateListItem should work if the ``sharedLockId`` could be provided as a parameter.

## Functionality
- Retrieve ``sharedLockId`` from the Site Pages module/component and show it.
- Three buttons to make requests to the validateUpdateListItem endpoint. 
  - First ValidateUpdateListItem: ``https://{SUB_DOMAIN}.sharepoint.com/sites/{SITE}/_api/web/GetList(@a1)/items(@a2)/ValidateUpdateListItem()?@a1='/sites/{SITE}/SitePages'&@a2='{LIST_ITEM_ID}'``
  - Second ValidateUpdateListItem: ``https://{SUB_DOMAIN}.sharepoint.com/sites/{SITE}/_api/web/lists('{LIST_ID}')/items({LIST_ITEM_ID})/validateupdatelistitem``
  - PnPJs 3.22 ValidateUpdateListItem: uses PnPJs to make the requests, but the endpoint is the same as the Second ValidateUpdateListItem button.
    - Since it is not possible to provide the ``sharedLockId`` as a parameter, this request will always fail.
- A toggle to enable/disable the ``sharedLockId`` in the request body.
  - If disabled, the requests made by the first and second buttons will also fail.

## Requirements
- Node version 18.20.4

## Minimal Path to Awesome
- Clone this repository
- Ensure that you are at the solution folder
- in the command-line run:
  - ``npm install``
  - ``npm run serve``
  - Open a browser and go to a SharePoint page.
  - Append the following query string to the URL: ``?debug=true&noredir=true&debugManifestsFile=https://localhost:4321/temp/manifests.js``
  - Edit the page.
  - Add the Cavuli web part to the page.

## References
- [Set up your SharePoint Framework development environment](https://learn.microsoft.com/en-us/sharepoint/dev/spfx/set-up-your-development-environment)
- [Getting started with SharePoint Framework](https://docs.microsoft.com/en-us/sharepoint/dev/spfx/set-up-your-developer-tenant)

## Disclaimer
**THIS CODE IS PROVIDED _AS IS_ WITHOUT WARRANTY OF ANY KIND, EITHER EXPRESS OR IMPLIED, INCLUDING ANY IMPLIED WARRANTIES OF FITNESS FOR A PARTICULAR PURPOSE, MERCHANTABILITY, OR NON-INFRINGEMENT.**

import { SPHttpClient, ISPHttpClientOptions, SPHttpClientResponse } from '@microsoft/sp-http';
import { WebPartContext } from '@microsoft/sp-webpart-base';
import { SPFx, spfi } from '@pnp/sp';
import '@pnp/sp/webs';
import '@pnp/sp/lists';
import '@pnp/sp/items';

export async function validateUpdateListItem(endpointUrl: string,context: WebPartContext, sharedLockId: string | null): Promise<string> {
  try {
    const body = {
      bNewDocumentUpdate: false,
      checkInComment: null,
      sharedLockId: sharedLockId,
      formValues: [
        {
          FieldName: 'Title',
          FieldValue: new Date().toString()
        }
      ]
    }

    const optionsWithData: ISPHttpClientOptions = {
      headers: {
        'Accept': 'application/json;odata=verbose',
        'Content-Type': 'application/json;odata=verbose',
        'OData-Version': '',
      },
      body: JSON.stringify(body)
    }

    const client = context.spHttpClient;

    const response: SPHttpClientResponse = await client.post(endpointUrl,
      SPHttpClient.configurations.v1, optionsWithData);
    const responseJson = await response.json();
    const json = responseJson?.d ?? responseJson;
    const result = json?.ValidateUpdateListItem?.results[0] ?? json.error;
    return JSON.stringify(result);
  } catch (error) {
    return JSON.stringify(error);

  }
}

export async function pnpValidateUpdateListItem(context: WebPartContext, sharedLockId: string | null): Promise<string> {

  const listId = context.pageContext.list?.id;
  const itemId = context.pageContext.listItem?.id;

  if (listId && itemId) {
    const sp = spfi().using(SPFx(context));
    try {
      const result = await sp.web.lists.getById(listId.toString()).items.getById(itemId).validateUpdateListItem(

        [
          {
            FieldName: 'Title',
            FieldValue: new Date().toString()
          }
        ]
        ,
        false
      );

      return JSON.stringify(result[0]);
    } catch (error) {
      return error.toString();
    }
  }

  return 'ListId or ListItemId is null';
}

import * as React from 'react';
import { DefaultButton, Stack, Text, Toggle } from '@fluentui/react';
import { WebPartContext } from '@microsoft/sp-webpart-base';
import { DisplayMode } from '@microsoft/sp-core-library';
import { pnpValidateUpdateListItem, validateUpdateListItem } from '../Functions';
import { useSharedLockId } from '../UseSharedLockId';

export interface ICavuliProps {
  displayMode: DisplayMode;
  context: WebPartContext;
}

export function Cavuli(props: ICavuliProps): JSX.Element {
  const {
    context,
    displayMode
  } = props;

  const [newResponse, setNewResponse] = React.useState<string>('');
  const [sendSharedLockId, setSendSharedLockId] = React.useState<boolean>(true);

  const sharedLockId = useSharedLockId();

  const _onClick = async (func: 'first' | 'second' | 'pnp'): Promise<void> => {
    if (sharedLockId !== null) {
      setNewResponse('loading...');
      let result = '';
      const sharedLockIdToSend = sendSharedLockId ? sharedLockId : null;

      if (func === 'first') {
        // https://{SUB_DOMAIN}.sharepoint.com/sites/{SITE}/_api/web/GetList(@a1)/items(@a2)/ValidateUpdateListItem()?@a1='/sites/{SITE}/SitePages'&@a2='{LIST_ITEM_ID}'
        const endpointUrl = `${context.pageContext.web.absoluteUrl}/_api/web/GetList(@a1)/items(@a2)/ValidateUpdateListItem()?@a1='${context.pageContext.web.serverRelativeUrl}/SitePages'&@a2='${context.pageContext.listItem?.id}'`;
        result = await validateUpdateListItem(endpointUrl, props.context, sharedLockIdToSend);

      } else if (func === 'second') {
        // https://{SUB_DOMAIN}.sharepoint.com/sites/{SITE}/_api/web/lists('{LIST_ID}')/items({LIST_ITEM_ID})/validateupdatelistitem
        const endpointUrl = `${context.pageContext.web.absoluteUrl}/_api/web/lists('${context.pageContext.list?.id.toString()}')/items(${context.pageContext.listItem?.id})/validateupdatelistitem`;
        result = await validateUpdateListItem(endpointUrl, props.context, sharedLockIdToSend);

      } else if (func === 'pnp') {
        result = await pnpValidateUpdateListItem(props.context, sharedLockIdToSend);
      }

      setNewResponse(result);
    }
  }

  return (
    <Stack tokens={{ childrenGap: '1rem' }}>
      <Stack horizontalAlign='center' tokens={{ childrenGap: '5px' }}>
        <Text variant='xLarge'>Co-authoring validateUpdateListItem</Text>
        {displayMode === DisplayMode.Read && <Text variant='medium'>SharedLockId is only available in Edit mode.</Text>}
        <Text>SharedLockId: {sharedLockId}</Text>
      </Stack >
      {displayMode === DisplayMode.Edit &&
        <>
          <Stack horizontalAlign='center'>
            <Text>Response:</Text>
            <Text>{newResponse}</Text>
          </Stack>
          <Stack>
            <Toggle
              label='Send SharedLockId to endpoints. If disabled, all requests should fail with 423 Locked.'
              onText='Send SharedLockId.'
              offText='Send null instead.'
              checked={sendSharedLockId}
              onChange={() => setSendSharedLockId(!sendSharedLockId)}
            />
          </Stack>
          <Stack horizontal horizontalAlign='space-evenly'>
            <Stack>
              <DefaultButton onClick={() => _onClick('first')}>First ValidateUpdateListItem</DefaultButton>
            </Stack>
            <Stack>
              <DefaultButton onClick={() => _onClick('second')}>Second ValidateUpdateListItem</DefaultButton>
            </Stack>
            <Stack>
              <DefaultButton onClick={() => _onClick('pnp')}>PnPJs 3.22 ValidateUpdateListItem</DefaultButton>
            </Stack>
          </Stack>
        </>
      }

    </Stack >
  );

}

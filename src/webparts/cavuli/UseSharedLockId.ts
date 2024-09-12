import * as React from "react";
import { SPComponentLoader } from '@microsoft/sp-loader';
import { DisplayMode } from "@microsoft/sp-core-library";

export function useSharedLockId(displayMode: DisplayMode): string | null {

  const [sharedLockId, setSharedLockId] = React.useState<string | null>(null);

  // This can be remade into a class function and run from CavuliWebPart.ts instead
  const getSharedLockId = React.useCallback(async () => {

    // Is the use of the module/component below unsupported?

    // ID refers to SPPages module/component fetched from here: \16\TEMPLATE\LAYOUTS\Next\spclient\b6917cb1-93a0-4b97-a84d-7cf49975d4ec.json
    const sitePagesComponent = await SPComponentLoader.loadComponentById<any>("b6917cb1-93a0-4b97-a84d-7cf49975d4ec"); // This is the GUID for the Site Pages feature

    if (sitePagesComponent?.PageStore?.fields) {
      const sharedLockId = sitePagesComponent.PageStore.fields.SharedLockId;
      setSharedLockId(sharedLockId);
    }
  }, []);

  React.useEffect(() => {
    if (displayMode === DisplayMode.Edit) {
      // eslint-disable-next-line @typescript-eslint/no-floating-promises
      getSharedLockId();
    }
  }, [displayMode]);

  return sharedLockId;
}

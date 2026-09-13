import { Config } from '../../../common/config/Config';
import { Consts } from '../../../common/consts/Consts';
import { PNPService } from '../../../common/services/PNPService';

export const navigateToDatasheet = (
  entityType: string,
  itemId: number
): void => {
  const { SITE } = Config;
  const { DATASHEET_URL } = Consts.Pages;

  if (!entityType || !itemId) {
    return;
  }

  const url = `${SITE.ROOT_WEB_URL}/${DATASHEET_URL}?config=${encodeURIComponent(entityType)}&itemId=${itemId}`;
  window.location.href = url;
};

export const createUniqueFolder = async (
  pnpService: PNPService,
  parentFolderUrl: string,
  baseName: string
): Promise<string> => {
  let folderName = baseName;
  let attempt = 0;

  while (true) {
    const targetFolderUrl = `${parentFolderUrl}/${folderName}`;
    // Check if the folder exists
    const targetFolder = await pnpService.getFolder(targetFolderUrl).select('Exists')();
    if (targetFolder.Exists) {
      attempt++;
      folderName = `${baseName}-${attempt}`;
    } else {
      await pnpService.getFolder(parentFolderUrl).folders.addUsingPath(folderName);
      return targetFolderUrl;
    }
  }
};
export const createFolder = async (
  pnpService: PNPService,
  parentFolderUrl: string,
  baseName: string
): Promise<string> => {
  const folderName = baseName;
  // Create the folder
  const targetFolderUrl = `${parentFolderUrl}/${folderName}`;
  await pnpService.getFolder(parentFolderUrl).folders.addUsingPath(folderName);
  return targetFolderUrl;


};
export const checkFolderExists = async (
  pnpService: PNPService,
  parentFolderUrl: string,
  baseName: string
): Promise<boolean> => {
  const folderName = baseName;
  const targetFolderUrl = `${parentFolderUrl}/${folderName}`;
  // Check if the folder exists
  const targetFolder = await pnpService.getFolder(targetFolderUrl).select('Exists')();
  return targetFolder.Exists;
};
export const datepickerstyles = {
  root: {
    fontFamily: 'inherit',
    color: 'inherit',
    width: '100%', // Force full width to prevent layout shifts
  },
  icon: {
    pointerEvents: 'none'
  },
  statusMessage: {
    fontWeight: 'normal'
  }
};

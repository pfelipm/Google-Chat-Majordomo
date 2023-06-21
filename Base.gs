/**
 * Google Chat Majordomo is a simple Apps Script automation based on spreadsheets, 
 * created for educational purposes to demonstrate how Google Chat API user authentication
 * methods can be leveraged. This tool implements a workflow that allows administrators
 * to review and approve requests for membership in different Google Chat spaces, collected
 * with Google Forms.
 * 
 * Copyright (C) Pablo Felip (@pfelipm) v1.0 JUN 2023.
 * Distributed under a GNU GPL v3 licence.
 *   
 * @OnlyCurrentDoc
 */

const PARAMS = {
  version: 'Version 1.0 (june 2023)',
  appName: 'Google Chat Majordomo',
  icon: '💢',
  toastTitle: '💢 Chat Majordomo says:',
  endpoints: {
    listSpaces: 'https://chat.googleapis.com/v1/spaces',
    spacesMembersCreate: 'https://chat.googleapis.com/v1'
  },
  sheets: {
    review: {
      name: 'Application review', dataRow: 3, colTimeStamp: 1, colEmail: 2, colSpace: 3, colCheck: 5, colLog: 6
    },
    settings: {
      name: 'Settings',
      formSpaceTable: 'A3:B', colFormItem: 1, colFormSpaceName: 2,
      spaceTable: 'D3:F', colSpaceName: 1, colSpaceId: 3,
    },
  },
  buttons: {
    leds: { process: 'H1', reload: 'C5' },
    status: { on: '🟢', off: '⚪' }
  },
  chatSpaceDescriptionMaxLength: 60,
};

/**
 * Builds custom menu.
 */
function onOpen() {

  const ui = SpreadsheetApp.getUi();
  ui.createMenu(`${PARAMS.icon} ${PARAMS.appName}`)
    .addItem(`💡 About ${PARAMS.appName}`, 'm_about')
    .addToUi();

}

/**
 * Shows the about this app dialog.
 */
function m_about() {

  const panel = HtmlService.createTemplateFromFile('About');
  panel.version = PARAMS.version;
  panel.appName = PARAMS.appName;
  SpreadsheetApp.getUi().showModalDialog(panel.evaluate().setWidth(450).setHeight(350), `💡 What is ${PARAMS.appName}?`);

}
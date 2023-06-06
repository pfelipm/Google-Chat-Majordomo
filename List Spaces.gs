/**
 * Gets the list of spaces the current user has access to and writes
 * names, descriptions and IDs in the spreadsheet.
 */
function listSpaces() {

  // Prevents concurrent runs
  const lock = LockService.getScriptLock()
  if (lock.tryLock(0)) {

    const ss = SpreadsheetApp.getActive()
    const s = ss.getSheetByName(PARAMS.sheets.settings.name);
    let result = [];
    let pageToken;

    // Signals start of process
    s.getRange(PARAMS.buttons.leds.reload).setValue(PARAMS.buttons.status.off);
    SpreadsheetApp.flush();
    ss.toast('Updating list of Chat spaces...', PARAMS.toastTitle, -1);

    // Pings the Chat API to get spaces using API
    do {
      const response = UrlFetchApp.fetch(
        encodeURI(`${PARAMS.endpoints.listSpaces}?filter=spaceType="SPACE"${pageToken ? `&pageToken=${pageToken}` : ''}`),
        {
          method: 'GET',
          muteHttpExceptions: true,
          headers: { Authorization: 'Bearer ' + ScriptApp.getOAuthToken() }
        }
      )
      if (response.getResponseCode() == 200) {
        const payload = JSON.parse(response.getContentText());
        result = result.concat(payload.spaces);
        pageToken = payload.nextPageToken;
      } else return response.getResponseCode();
    } while (pageToken);

    // Gets only desired fields from reponse using nested destructuring and sort spaces by name
    const spacesData = result.map(space => {
      const { name, displayName, spaceDetails: { description = 'No description available' } = '' } = space;
      return [
        displayName,
        description.length > PARAMS.chatSpaceDescriptionMaxLength -3 ? `${description.slice(0, 60)}...` : description,
        name
      ];
    }).sort((space1, space2) => space1[0].localeCompare(space2[0]));

    // Writes data in table of spaces
    s.getRange(PARAMS.sheets.settings.spaceTable)
      .clearContent()
      .offset(0, 0, spacesData.length, spacesData[0].length).setValues(spacesData);

    // Signals end of process
    ss.toast('Done!', PARAMS.toastTitle);
    s.getRange(PARAMS.buttons.leds.reload).setValue(PARAMS.buttons.status.on);

    // Unnecessary, but recommended
    SpreadsheetApp.flush();
    lock.releaseLock();

  }

}

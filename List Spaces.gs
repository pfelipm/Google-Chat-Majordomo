/**
 * Gets the list of spaces the current user has access to and writes
 * names, descriptions and IDs in the spreadsheet.
 */
function listSpaces() {

  // Prevents concurrent runs
  const lock = LockService.getScriptLock();
  if (lock.tryLock(0)) {

    const ss = SpreadsheetApp.getActive();
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
        // Double quotes (") are not valid in the URI, either replace with %22 or use encodeURI()
        encodeURI(`${PARAMS.endpoints.listSpaces}?filter=spaceType="SPACE"${pageToken ? `&pageToken=${pageToken}` : ''}`),
        {
          method: 'GET',
          muteHttpExceptions: true,
          headers: { Authorization: 'Bearer ' + ScriptApp.getOAuthToken() }
        }
      );
      if (response.getResponseCode() == 200) {
        const payload = JSON.parse(response.getContentText());
        result = result.concat(payload.spaces);
        pageToken = payload.nextPageToken;
      } else ss.toast(`Error ${response.getResponseCode()}.`, PARAMS.toastTitle);
    } while (pageToken);

    // Gets only the desired fields from the result array using nested destructuring and sorts spaces by name
    const spacesData = result.map(space => {
      const { name, displayName, spaceDetails: { description = 'No description available' } = '' } = space;
      return [
        displayName,
        description.length > PARAMS.chatSpaceDescriptionMaxLength - 3 ? `${description.slice(0, 60)}...` : description,
        // Space ID, actually
        name
      ];
    // Sorts by displayName
    }).sort((space1, space2) => space1[0].localeCompare(space2[0]));

    // Writes data, if any, in table (directory) of spaces
    if (spacesData.length > 0) {
      s.getRange(PARAMS.sheets.settings.spaceTable)
        .clearContent()
        .offset(0, 0, spacesData.length, spacesData[0].length).setValues(spacesData);

      // Signals successful end of process
      ss.toast('Done!', PARAMS.toastTitle);
    }

    // Restores the green indicator circle
    s.getRange(PARAMS.buttons.leds.reload).setValue(PARAMS.buttons.status.on);
    // Unnecessary, but recommended
    SpreadsheetApp.flush();    
    lock.releaseLock();

  }

}
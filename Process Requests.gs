/**
 * Adds users to Chat spaces, if granted access.
 */
function processRequests() {

  // Prevents concurrent runs
  const lock = LockService.getScriptLock()
  if (lock.tryLock(0)) {

    const ss = SpreadsheetApp.getActive()
    const s = ss.getSheetByName(PARAMS.sheets.review.name);
    let usersAdded = 0;

    const requests = s.getDataRange().getValues().slice(PARAMS.sheets.review.dataRow - 1);
    const numRequests = requests.filter(row => row[PARAMS.sheets.review.colTimeStamp - 1] && row[PARAMS.sheets.review.colCheck - 1]).length;

    if (numRequests == 0) ss.toast('No pending approved requests!', PARAMS.toastTitle);
    else {

      // Let's add users to spaces

      // Signals start of process
      s.getRange(PARAMS.buttons.leds.process).setValue(PARAMS.buttons.status.off);
      SpreadsheetApp.flush();
      ss.toast('Adding approved users to Chat spaces ...', PARAMS.toastTitle, -1);

      // Reads form item to Chat space name table
      const settingsSheet = SpreadsheetApp.getActive().getSheetByName(PARAMS.sheets.settings.name);
      const formSpaceTableValues = settingsSheet.getRange(PARAMS.sheets.settings.formSpaceTable).getValues();

      // Reads Chat space name to Chat space ID table
      const spaceTableValues = settingsSheet.getRange(PARAMS.sheets.settings.spaceTable).getValues();

      // Loops over the requests and filter out those that should not be processed
      requests.forEach((request, index, array) => {

        const email = request[PARAMS.sheets.review.colEmail - 1];
        const chatSpaceItem = request[PARAMS.sheets.review.colSpace - 1];
        const check = request[PARAMS.sheets.review.colCheck - 1];
        if (check) {

          const chatSpaceName = formSpaceTableValues
            .find(item => item[PARAMS.sheets.settings.colFormItem - 1] == chatSpaceItem)?.[PARAMS.sheets.settings.colFormSpaceName - 1];
          const chatSpaceId = spaceTableValues
            .find(chatSpace => chatSpace[PARAMS.sheets.settings.colSpaceName - 1] == chatSpaceName)?.[PARAMS.sheets.settings.colSpaceId - 1];
          if (!chatSpaceId) array[index][PARAMS.sheets.review.colLog - 1] = `Can't find space!`;
          else {

            let userId;
            try {

              // Gets user's id using the Directory API
              userId = AdminDirectory.Users.get(email, { projection: 'BASIC', viewType: 'domain_public' }).id;
              
              console.info(userId);

              // Adds user to space

              const response = UrlFetchApp.fetch(
                `${PARAMS.endpoints.spacesMembersCreate}/${chatSpaceId}/members`,
                {
                  method: 'POST',
                  muteHttpExceptions: true,
                  headers: { Authorization: 'Bearer ' + ScriptApp.getOAuthToken() },
                  contentType: 'application/json',
                  payload: JSON.stringify({ member: { name: `users/${userId}`, type: 'HUMAN' } })                }
              );
              
              console.info(response.getResponseCode(), response.getContentText());
              
              if (response.getResponseCode() != 200) array[index][PARAMS.sheets.review.colLog - 1] = `Can't add user to space!`;
              else {
                // Uncheck request and write timestamp in log
                usersAdded++;
                array[index][PARAMS.sheets.review.colCheck -1] = false;
                array[index][PARAMS.sheets.review.colLog - 1] =  Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'dd/MM/yyyy HH:mm');
              }

            } catch (e) {
              // This can fail is requesting user is external to the domain
              console.info('User not found!');
              array[index][PARAMS.sheets.review.colLog - 1] = `Can't find user!`;
            }

          }
        }

      });

      // Update access grating checkboxes in sheet
      s.getRange(PARAMS.sheets.review.dataRow, PARAMS.sheets.review.colCheck, requests.length, 1)
        .setValues(requests.map(request => [request[PARAMS.sheets.review.colCheck - 1]]));

      // Write log to sheet, could be done in one step together with checkboxes, but columns could not be next to each other
      s.getRange(PARAMS.sheets.review.dataRow, PARAMS.sheets.review.colLog, requests.length, 1)
        .setValues(requests.map(request => [request[PARAMS.sheets.review.colLog - 1]]));

      // Signals end of process
      ss.toast(`Done! (added ${usersAdded} 👤)`, PARAMS.toastTitle);
      s.getRange(PARAMS.buttons.leds.process).setValue(PARAMS.buttons.status.on);

      // Unnecessary, but recommended
      SpreadsheetApp.flush();
      lock.releaseLock();

    }

  }

}

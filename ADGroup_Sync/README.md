# AD Group Sync

This script allows synchronization of Active Directory Groups to Webex Spaces

### Deployment Steps
1. Register a Bot at [Webex Developers](https://developer.webex.com/my-apps)
2. Download PS1 and CSV Files, intention is to run on Windows System via Task Scheduler
    - **Note:** Script must be run by an account with AD Read Permissions
3. Edit Script and update the following variables
    - WebexAuth (required) - Webex Authentication token for Bot (registered in Step 1)
    - ReportId (optional) - UserId or RoomId to send error status updates
    - ExemptUsers (optional) - Array of users to be excluded from the Sync **removal** process
4. Update CSV File with AD Group to Webex Space details
    - **Note:** RoomId is required, this can be easily obtained from the [Developer API](https://developer.webex.com/docs/api/v1/rooms/list-rooms) or by adding `astronaut@webex.bot` to the space
5. Add the bot registered above to applicable Webex Spaces
6. Run Script! You can use the argument -Debug to get better visibility of the process
7. Use Windows Task Scheduler to automate exection of this script on a regular basis to keep Webex in sync

## Support

In case you've found a bug, please [open an issue on GitHub](../../issues).

## Disclamer

This script is NOT guaranteed to be bug free and production quality.
# AD Group Sync

This script allows synchronization of Active Directory Groups to Webex Spaces, Teams or Individual Team Spaces

**Note:** Sync is performed from AD to Webex, not in reverse.

## Usage Notes

You can synchronize Team Membership by mapping to the "General" space for the respective team.

**Note:** Removal from the Team **WILL** result in removal from all Team Spaces!

If you are performing sync on individual team spaces, please order your CSV file to process Teams (General Spaces) first to ensure that users who are meant to be part of an individual team spaces are re-added if they get removed from the Team.

## Deployment Steps
**NOTE:** This now requires Powershell 7+ to function, along with a supported RSAT version for Active Directory compatibility with PS7

1. Register a Bot at [Webex Developers](https://developer.webex.com/my-apps)
2. Download PS1 and CSV Files, intention is to run on Windows System via Task Scheduler
    - **Note:** Script must be run by an account with AD Read Permissions
3. Edit Script and update the following variables
    - WebexAuth (required) - Webex Authentication token for Bot (registered in Step 1)
    - ReportId (optional) - Person Id or Room Id to send error status updates
    - ExemptUsers (optional) - Array of users to be excluded from the Sync **removal** process
4. Update CSV File with AD Group to Webex Space details, these example methods can be used to get the Room Id
    - Using the [List Rooms](https://developer.webex.com/docs/api/v1/rooms/list-rooms) Developer API
    - Adding `astronaut@webex.bot` to the space (bot will leave and 1:1 you the Id)
    - 1:1 Message `astronaut@webex.bot`, with an @Mention of the Space name
5. Add the bot registered above to applicable Webex Spaces
6. Run Script! You can use the argument -Debug to get better visibility of the process
7. Use Windows Task Scheduler to automate execution of this script on a regular basis to keep Webex in sync

## Support

In case you've found a bug, please [open an issue on GitHub](../../../issues).

## Disclaimer

This script is NOT guaranteed to be bug free and production quality.

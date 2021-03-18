# Azure AD Connect ObjectId Sync Rule

This script to add a new Azure AD Connect Synchronization Rule to sync the Azure Object ID to on-prem AD

Webex Directory Connector (DirSync) can then sync this to the Webex Cloud (externalId attribute)

## Deployment 
### Automatic
1. Download and run the included Powershell script on the Azure AD Connect Server (Run as Administrator)

### Manual
1. Launch the Azure AD Connect Synchronization Rules Editor (Run as Administrator)
2. Add a new **Outbound** Rule
    #### **Description**
    - Name: Out to AD - User AAD ObjectId
    - Description: Maps AAD ObjectId to AD Attribute for Webex
    - Connected System: `<Your AD Domain>`
    - Connected System Object Type: user
    - Metaverse Object Type: person
    - Link Type: Join
    - Precedence: 148 (must be Unique, adjust as required)
    #### **Scoping filter**
    - Attribute: sourceAnchor
    - Operator: ISNOTNULL
    #### **Join rules**
    - Blank
    #### **Transformations**
    - FlowType: Expression
    - Target Attribute: `msDS-cloudExtensionAttribute1` (adjust as required)
    - Source: `Replace([cloudAnchor],"User_","")`
    - Apply Once: Unchecked
    - Merge Type: Update

## Support

In case you've found a bug, please [open an issue on GitHub](../../../issues).

## Disclamer

This script is NOT guaranteed to be bug free and production quality.
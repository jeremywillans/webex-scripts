New-ADSyncRule  `
-Name 'Out to AD - User ExternalId for Webex' `
-Identifier '7da2fd41-61d4-463e-9a0f-7b4457654bb2' `
-Description '' `
-Direction 'Outbound' `
-Precedence 148 `
-PrecedenceAfter '00000000-0000-0000-0000-000000000000' `
-PrecedenceBefore '00000000-0000-0000-0000-000000000000' `
-SourceObjectType 'person' `
-TargetObjectType 'user' `
-Connector '4a85c66f-18fc-4937-bad3-b282fbff6b2c' `
-LinkType 'Join' `
-SoftDeleteExpiryInterval 0 `
-ImmutableTag '' `
-OutVariable syncRule


Add-ADSyncAttributeFlowMapping  `
-SynchronizationRule $syncRule[0] `
-Destination 'msDS-cloudExtensionAttribute1' `
-FlowType 'Expression' `
-ValueMergeType 'Update' `
-Expression 'Replace([cloudAnchor],"User_","")' `
-OutVariable syncRule


New-Object  `
-TypeName 'Microsoft.IdentityManagement.PowerShell.ObjectModel.ScopeCondition' `
-ArgumentList 'sourceAnchor','','ISNOTNULL' `
-OutVariable condition0


Add-ADSyncScopeConditionGroup  `
-SynchronizationRule $syncRule[0] `
-ScopeConditions @($condition0[0]) `
-OutVariable syncRule


Add-ADSyncRule  `
-SynchronizationRule $syncRule[0]


Get-ADSyncRule  `
-Identifier '7da2fd41-61d4-463e-9a0f-7b4457654bb2'
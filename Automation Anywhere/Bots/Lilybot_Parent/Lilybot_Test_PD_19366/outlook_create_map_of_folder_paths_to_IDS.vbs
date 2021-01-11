Dim fso, outFile
Set fso = CreateObject("Scripting.FileSystemObject")
Set outFile = fso.CreateTextFile("output.txt", True)

'  This example requires the Chilkat API to have been previously unlocked.
'  See Global Unlock Sample for sample code.
'  11/24/2020

set http = CreateObject("Chilkat_9_5_0.Http")

'  Our folder path --> ID map will be stored in this hash table.
set folderMap = CreateObject("Chilkat_9_5_0.Hashtable")

'  Use your previously obtained access token here:
'  See the following examples for getting an access token:
'     Get Microsoft Graph OAuth2 Access Token (Azure AD v2.0 Endpoint).
'     Get Microsoft Graph OAuth2 Access Token (Azure AD Endpoint).
'     Refresh Access Token (Azure AD v2.0 Endpoint).
'     Refresh Access Token (Azure AD Endpoint).

http.AuthToken = "MICROSOFT_GRAPH_ACCESS_TOKEN"

set sbResponse = CreateObject("Chilkat_9_5_0.StringBuilder")

'  Begin by getting the top-level folders.
http.ClearUrlVars 
success = http.SetUrlVar("userPrincipalName","chilkatsoft@outlook.com")
success = http.QuickGetSb("https://graph.microsoft.com/v1.0/users/{$userPrincipalName}/mailFolders",sbResponse)
If ((success <> 1) And (http.LastStatus = 0)) Then
    outFile.WriteLine(http.LastErrorText)
    WScript.Quit
End If

set json = CreateObject("Chilkat_9_5_0.JsonObject")
success = json.LoadSb(sbResponse)
json.EmitCompact = 0

outFile.WriteLine("Status code = " & http.LastStatus)
If (http.LastStatus <> 200) Then
    outFile.WriteLine(json.Emit())
    outFile.WriteLine("Failed.")
End If

'  This is our queue/stack of unprocessed folder ID's
'  The recursive nature of this example is that we get the
'  child folders for each folder ID in the idQueue, which may
'  cause additional ID's to be added.  We continue  until the idQueue
'  is empty.
set idQueue = CreateObject("Chilkat_9_5_0.StringArray")

set sbFolderPath = CreateObject("Chilkat_9_5_0.StringBuilder")
set sbQueueEntry = CreateObject("Chilkat_9_5_0.StringBuilder")

'  Prime the map and idQueue with the top-level folders.
i = 0
numFolders = json.SizeOfArray("value")
Do While i < numFolders
    json.I = i
    folderName = json.StringOf("value[i].displayName")
    folderId = json.StringOf("value[i].id")
    success = sbFolderPath.SetString("/")
    success = sbFolderPath.Append(folderName)
    folderPath = sbFolderPath.GetAsString()
    success = folderMap.AddStr(folderPath,folderId)
    outFile.WriteLine(folderPath & " --> " & folderId)

    '  Push the folder path + id onto the idQueue.
    sbQueueEntry.Clear 
    success = sbQueueEntry.SetNth(0,folderPath,"|",0,0)
    success = sbQueueEntry.SetNth(1,folderId,"|",0,0)
    success = idQueue.Append(sbQueueEntry.GetAsString())
    i = i + 1
Loop

'  Initial output:
'  /Archive --> AQMkADAwATM0MDAAMS1iNTcwLWI2NTEtMDACLTAwCgAuAAADsVyfxjDU406Ic4X7ill8xAEA5_vF7TKKdE6bGCRqXyl2PQAAAG8XunwAAAA=
'  /Deleted Items --> AQMkADAwATM0MDAAMS1iNTcwLWI2NTEtMDACLTAwCgAuAAADsVyfxjDU406Ic4X7ill8xAEA5_vF7TKKdE6bGCRqXyl2PQAAAgEKAAAA
'  /Drafts --> AQMkADAwATM0MDAAMS1iNTcwLWI2NTEtMDACLTAwCgAuAAADsVyfxjDU406Ic4X7ill8xAEA5_vF7TKKdE6bGCRqXyl2PQAAAgEPAAAA
'  /Inbox --> AQMkADAwATM0MDAAMS1iNTcwLWI2NTEtMDACLTAwCgAuAAADsVyfxjDU406Ic4X7ill8xAEA5_vF7TKKdE6bGCRqXyl2PQAAAgEMAAAA
'  /Junk Email --> AQMkADAwATM0MDAAMS1iNTcwLWI2NTEtMDACLTAwCgAuAAADsVyfxjDU406Ic4X7ill8xAEA5_vF7TKKdE6bGCRqXyl2PQAAAgEiAAAA
'  /Outbox --> AQMkADAwATM0MDAAMS1iNTcwLWI2NTEtMDACLTAwCgAuAAADsVyfxjDU406Ic4X7ill8xAEA5_vF7TKKdE6bGCRqXyl2PQAAAgELAAAA
'  /Sent Items --> AQMkADAwATM0MDAAMS1iNTcwLWI2NTEtMDACLTAwCgAuAAADsVyfxjDU406Ic4X7ill8xAEA5_vF7TKKdE6bGCRqXyl2PQAAAgEJAAAA
' 

'  Process the idQueue until it becomes empty.  This is the recursive loop.

Do While idQueue.Length > 0
    success = sbQueueEntry.SetString(idQueue.GetString(0))
    success = idQueue.RemoveAt(0)

    parentFolderPath = sbQueueEntry.GetNth(0,"|",0,0)
    parentFolderId = sbQueueEntry.GetNth(1,"|",0,0)

    success = http.SetUrlVar("id",parentFolderId)
    success = http.QuickGetSb("https://graph.microsoft.com/v1.0/users/{$userPrincipalName}/mailFolders/{$id}/childFolders",sbResponse)
    If ((success <> 1) And (http.LastStatus = 0)) Then
        outFile.WriteLine(http.LastErrorText)
        WScript.Quit
    End If

    success = json.LoadSb(sbResponse)
    If (http.LastStatus <> 200) Then
        outFile.WriteLine("Status code = " & http.LastStatus)
        outFile.WriteLine(json.Emit())
        outFile.WriteLine("Failed.")
    End If

    i = 0
    numFolders = json.SizeOfArray("value")
    Do While i < numFolders
        json.I = i
        folderName = json.StringOf("value[i].displayName")
        folderId = json.StringOf("value[i].id")
        success = sbFolderPath.SetString(parentFolderPath)
        success = sbFolderPath.Append("/")
        success = sbFolderPath.Append(folderName)
        folderPath = sbFolderPath.GetAsString()
        success = folderMap.AddStr(folderPath,folderId)
        outFile.WriteLine(folderPath & " --> " & folderId)

        '  Push the folder path + id onto the idQueue.
        sbQueueEntry.Clear 
        success = sbQueueEntry.SetNth(0,folderPath,"|",0,0)
        success = sbQueueEntry.SetNth(1,folderId,"|",0,0)
        success = idQueue.Append(sbQueueEntry.GetAsString())
        i = i + 1
    Loop

Loop

'  The hash table of mail folder paths --> ID's can be persisted to XML and saved to a file or database (or anywhere..)
set sbFolderMapXml = CreateObject("Chilkat_9_5_0.StringBuilder")
success = folderMap.ToXmlSb(sbFolderMapXml)
success = sbFolderMapXml.WriteFile("qa_data/outlook/folderMap.xml","utf-8",0)

'  The hash table can be restored from the serialized XML like this:
set ht2 = CreateObject("Chilkat_9_5_0.Hashtable")
set sb2 = CreateObject("Chilkat_9_5_0.StringBuilder")
success = sb2.LoadFile("qa_data/outlook/folderMap.xml","utf-8")
success = ht2.AddFromXmlSb(sb2)

'  What's the ID for the folder "/Inbox/abc/subFolderA" ?
outFile.WriteLine("id for /Inbox/abc/subFolderA = " & ht2.LookupStr("/Inbox/abc/subFolderA"))

'  Final output:

'  /Archive --> AQMkADAwATM0MDAAMS1iNTcwLWI2NTEtMDACLTAwCgAuAAADsVyfxjDU406Ic4X7ill8xAEA5_vF7TKKdE6bGCRqXyl2PQAAAG8XunwAAAA=
'  /Deleted Items --> AQMkADAwATM0MDAAMS1iNTcwLWI2NTEtMDACLTAwCgAuAAADsVyfxjDU406Ic4X7ill8xAEA5_vF7TKKdE6bGCRqXyl2PQAAAgEKAAAA
'  /Drafts --> AQMkADAwATM0MDAAMS1iNTcwLWI2NTEtMDACLTAwCgAuAAADsVyfxjDU406Ic4X7ill8xAEA5_vF7TKKdE6bGCRqXyl2PQAAAgEPAAAA
'  /Inbox --> AQMkADAwATM0MDAAMS1iNTcwLWI2NTEtMDACLTAwCgAuAAADsVyfxjDU406Ic4X7ill8xAEA5_vF7TKKdE6bGCRqXyl2PQAAAgEMAAAA
'  /Junk Email --> AQMkADAwATM0MDAAMS1iNTcwLWI2NTEtMDACLTAwCgAuAAADsVyfxjDU406Ic4X7ill8xAEA5_vF7TKKdE6bGCRqXyl2PQAAAgEiAAAA
'  /Outbox --> AQMkADAwATM0MDAAMS1iNTcwLWI2NTEtMDACLTAwCgAuAAADsVyfxjDU406Ic4X7ill8xAEA5_vF7TKKdE6bGCRqXyl2PQAAAgELAAAA
'  /Sent Items --> AQMkADAwATM0MDAAMS1iNTcwLWI2NTEtMDACLTAwCgAuAAADsVyfxjDU406Ic4X7ill8xAEA5_vF7TKKdE6bGCRqXyl2PQAAAgEJAAAA
'  /Inbox/abc --> AQMkADAwATM0MDAAMS1iNTcwLWI2NTEtMDACLTAwCgAuAAADsVyfxjDU406Ic4X7ill8xAEA5_vF7TKKdE6bGCRqXyl2PQAAAL8huv8AAAA=
'  /Inbox/xyz --> AQMkADAwATM0MDAAMS1iNTcwLWI2NTEtMDACLTAwCgAuAAADsVyfxjDU406Ic4X7ill8xAEA5_vF7TKKdE6bGCRqXyl2PQAAAL8huwEAAAA=
'  /Inbox/abc/subFolderA --> AQMkADAwATM0MDAAMS1iNTcwLWI2NTEtMDACLTAwCgAuAAADsVyfxjDU406Ic4X7ill8xAEA5_vF7TKKdE6bGCRqXyl2PQAAAL8huwAAAQ==
'  /Inbox/abc/subFolderB --> AQMkADAwATM0MDAAMS1iNTcwLWI2NTEtMDACLTAwCgAuAAADsVyfxjDU406Ic4X7ill8xAEA5_vF7TKKdE6bGCRqXyl2PQAAAL8huwMAAAA=
'  /Inbox/abc/subFolderA/a --> AQMkADAwATM0MDAAMS1iNTcwLWI2NTEtMDACLTAwCgAuAAADsVyfxjDU406Ic4X7ill8xAEA5_vF7TKKdE6bGCRqXyl2PQAAAL8huwIAAAA=
' 

'  ------------------------------------------------------------------------------------------------------
'  This example applies to: Exchange Online | Office 365 | Hotmail.com | Live.com | MSN.com | Outlook.com | Passport.com
' 
'  The Microsoft Graph Outlook Mail API lets you read, create, and send messages and attachments,
'  view and respond to event messages, and manage folders that are secured by Azure Active Directory
'  in Office 365. It also provides the same functionality in Microsoft accounts specifically
'  in these domains: Hotmail.com, Live.com, MSN.com, Outlook.com, and Passport.com.

outFile.Close

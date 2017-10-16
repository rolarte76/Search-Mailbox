$Menu = @"
=========================================================================
1. Estimate Results Only for a  mailbox
2. Estimate Results Only for a list of mailboxes
3. Search a mailbox with sender's email address and subject name
4. Search a mailbox with sender's email address only
5. Search a mailbox with subject name only
6. Search Attachement name for a list of mailboxes
7. Search Using a date range & subject
8. Exit, or Press CTRL + C
=========================================================================
Enter # from the options above 1 - 7
"@

$Number = Read-Host $Menu


switch ($number) {

1 {

$Sender = Read-Host "Type sender's email address"
            $SearchUsermbx = Read-host "Type email address of mailbox being searched"
            $SubjectName = Read-Host "Type Subject Name"
            $DateRecieved = Read-Host "Type date received in the following format: mm/dd/yyyy" 
            #Search-Mailbox -identity $SearchUsermbx -SearchQuery "(Received:$DateRecieved) AND (From:$Sender) AND (Subject:$SubjectName)" -LogLevel full
            Search-Mailbox -identity $SearchUsermbx -SearchQuery "(Received:$DateRecieved) AND (From:$Sender) AND (Subject:$SubjectName)" -EstimateResultOnly
            
}


2 {

 $Sender = Read-Host "Type sender's email address"
            $SubjectName = Read-Host "Type Subject Name"
            $DateRecieved = Read-Host "Type date received in the following format: mm/dd/yyyy"
            Write-host "Ensure the source csv file has a heading called:'recipients'" -BackgroundColor Green
            Write-host "Enter the list of email addresses in a column below the 'recipients' heading" -BackgroundColor Green
            Write-Host "Sample SHARED path: \\EXCH16\SearchMailboxShared\ListUsers.csv" -BackgroundColor Green
            $ImportListUsers = Read-Host "Enter the full SHARED Path of the source list of mailboxes to be searched" 
            Import-csv $ImportListUsers | Foreach {Search-Mailbox -Identity $_.recipients -SearchQuery "(Received:$DateRecieved) AND (From:$Sender) AND (Subject:$SubjectName)" -EstimateResultOnly}
            Write-Host "After validating the output, you may proceed with the removal process or removal tool." -BackgroundColor Green; break

}

3 {
            $Sender = Read-Host "Type sender's email address"
            Write-host "Do not use a subject such as: [':RE']..." -ForegroundColor DarkYellow
            Write-host "It's best not to select a subject, as it may delete unwanted messages. Press CTRL+C and start over." -ForegroundColor DarkYellow
            $SearchUsermbx = Read-host "Type email address of mailbox being searched"
            $SubjectName = Read-Host "Type Subject Name"
            $DateRecieved = Read-Host "Type date received in the following format: mm/dd/yyyy"
            $TargetMailbox = Read-Host "Type the alias, or email address of the target mailbox where messages will be copied to"
            $Target_Folder = Read-Host "Type Folder Name where the message(s) will be moved. Do not use quotes."
            Write-Output "Started: $(Get-Date -format T)"
            Search-Mailbox -Identity $SearchUsermbx -SearchQuery "(Received:$DateRecieved) AND (From:$Sender) AND (Subject:$SubjectName)" -targetMailbox $TargetMailbox -TargetFolder $Target_Folder -LogLevel full | Out-File "C:\OutputReport.txt" -Append
            Write-Output "Ended: $(Get-Date -Format T)"
            Write-Host "Check the target mailbox' folder name provided:" $Target_Folder
            Write-Host "Please, review the output report found on path C:\OutputReport.txt. Once you have validated the information, you may proceed by removing the message items by utilizing the '-deletecontent -force'" -BackgroundColor Green; break
}

4 {
            $Sender = Read-Host "Type sender's email address"
            $SearchUsermbx = Read-host "Type email address of mailbox being searched"
            $DateRecieved = Read-Host "Type date received in the following format: mm/dd/yyyy"
            $TargetMailbox = Read-Host "Type the alias, or email address of the target mailbox where messages will be copied to"
            $Target_Folder = Read-Host "Type Folder Name where the message(s) will be moved. Do not use quotes."
            Write-Output "Started: $(Get-Date -format T)"
            Search-Mailbox -Identity $SearchUsermbx -SearchQuery "(Received:$DateRecieved) AND (From:$Sender)" -targetMailbox $TargetMailbox -TargetFolder $Target_Folder -LogLevel full | Out-File "C:\OutputReport.txt"
            Write-Output "Ended: $(Get-Date -Format T)"
            Write-Host "Check the target mailbox' folder name provided:" $Target_Folder
            Write-Host "Please, review the output report found on path C:\OutputReport.txt. Once you have validated the information, you may proceed by removing the message items by utilizing the '-deletecontent -force'" -BackgroundColor Green; break
}

5 {
            $SubjectName = Read-Host "Be aware if you will be using a subject such as ':RE' only, it is best not to select a subject, as it may delete unwanted messages. Press CTRL+C and start over. Type Subject Name"
            $SearchUsermbx = Read-host "Type email address of mailbox being searched"
            $DateRecieved = Read-Host "Type date received in the following format: mm/dd/yyyy"
            $TargetMailbox = Read-Host "Type the alias, or email address of the target mailbox where messages will be copied to"
            $Target_Folder = Read-Host "Type Folder Name where the message(s) will be moved. Do not use quotes."
            Write-Output "Started: $(Get-Date -format T)"
            Search-Mailbox -Identity $SearchUsermbx -SearchQuery "(Received:$DateRecieved) AND (Subject:$SubjectName)" -targetMailbox $TargetMailbox -TargetFolder $Target_Folder -LogLevel full | Out-File "C:\Temp\OutputReport.txt"
            Write-Output "Ended: $(Get-Date -Format T)"
            Write-Host "Check the target mailbox' folder name provided:" $Target_Folder
            Write-Host "Please, review the output report found on path C:\OutputReport.txt. Once you have validated the information, you may proceed by removing the message items by utilizing the '-deletecontent -force'" -BackgroundColor Green; break
}


6 {
        
            $Target_Folder = Read-Host "Type Folder Name where the message(s) will be moved. Do not use quotes."
            $FileAttachName = Read-Host "Type name of the attachement including the extension file type, example 'filename.txt'"
            $TargetMailbox = Read-Host "Type the alias, or email address of the target mailbox where messages will be copied to"
            Write-host "Ensure the source csv file has a heading called:'recipients'" -BackgroundColor Green
            Write-host "Enter the list of email addresses in a column below the 'recipients' heading" -BackgroundColor Green
            Write-Host "Sample SHARED path: \\EXCH16\SearchMailboxShared\ListUsers.csv" -BackgroundColor Green
            $ImportListUsers = Read-Host "Enter the full SHARED Path of the source list of mailboxes to be searched" 
            Import-csv $ImportListUsers | Foreach {Search-Mailbox -Identity $_.recipients -SearchQuery $FileAttachName -targetMailbox $TargetMailbox -TargetFolder $Target_Folder -LogLevel full | Out-File "C:\OutputReport.txt" -Append}
            Write-Host "Check the target mailbox' folder name provided:"
            Write-Host "Please, review the output report found on path C:\OutputReport.txt. Once you have validated the information, you may proceed by removing the message items by utilizing the '-deletecontent -force'" -BackgroundColor Green; break
}

7 {
  
            $SubjectName = Read-Host "Be aware if you will be using a subject such as ':RE' only, it is best not to select a subject, as it may delete unwanted messages. Press CTRL+C and start over. Type Subject Name"
            $SearchUsermbx = Read-host "Type email address of mailbox being searched"
            $DateRecieved = Read-Host "Type date range date received in the following format, use the dots on between: 'mm/dd/yyyy..mm/dd/yyyy'"                    
            $TargetMailbox = Read-Host "Type the alias, or email address of the target mailbox where messages will be copied to"
            $Target_Folder = Read-Host "Type Folder Name where the message(s) will be moved. Do not use quotes."
            Write-Output "Started: $(Get-Date -format T)"
            Search-Mailbox -Identity $SearchUsermbx -SearchQuery "(Received:$DateRecieved) AND (Subject:$SubjectName)" -targetMailbox $TargetMailbox -TargetFolder $Target_Folder -LogLevel full | Out-File "C:\OutputReport.txt"
            Write-Output "Ended: $(Get-Date -Format T)"
            Write-Host "Check the target mailbox' folder name provided:" $Target_Folder
            Write-Host "Please, review the output report found on path C:\OutputReport.txt. Once you have validated the information, you may proceed by removing the message items by utilizing the '-deletecontent -force'" -BackgroundColor Green; break
}

Default {"There were no other options selected"}
}

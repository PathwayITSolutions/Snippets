<#$User = "[USER]"
$Pass = $env:svcdscript | ConvertTo-SecureString -AsPlainText -ForegroundColor
$LiveCred = New-Object -TypeName System.Management.Automation.PSCredential -ArgumentList $User,$Pass

$SMTPFrom = "[NOREPLY]"
$SMTPTo = "[RECIPIENT]"
$Subject = "Forwarding mail rules to external"
$SMTPServer = "smtp.office365.com"
$SMTPPort = "587"#>
$Rundate = Get-Date -Format "MMM-yyyy"
$domains = Get-AcceptedDomain
$mailboxes = Get-Mailbox -ResultSize Unlimited
$RuleDb = New-Object System.Collections.ArrayList
 
foreach ($mailbox in $mailboxes) {
 
    $forwardingRules = $null
    Write-Host "Checking rules for $($mailbox.displayname) - $($mailbox.primarysmtpaddress)" -foregroundColor Green
    $rules = get-inboxrule -Mailbox $mailbox.primarysmtpaddress
     
    $forwardingRules = $rules | Where-Object { $_.forwardto -or $_.forwardasattachmentto }
 
    foreach ($rule in $forwardingRules) {
        $recipients = New-Object System.Collections.ArrayList
        $recipientsForwardTo = $rule.ForwardTo | Where-Object { $_ -match "SMTP" }
        $recipientsForwardAsAttachmentTo = $rule.ForwardAsAttachmentTo | Where-Object { $_ -match "SMTP" }

        if ($recipientsForwardTo) { $recipients.Add($recipientsForwardTo) | Out-Null }
        if ($recipientsForwardAsAttachmentTo) { $recipients.Add($recipientsForwardAsAttachmentTo) | Out-Null }
     
        $externalRecipients = New-Object System.Collections.ArrayList
 
        foreach ($recipient in $recipients) {
            $email = ($recipient -split "SMTP:")[1].Trim("]")
            $domain = ($email -split "@")[1]
 
            if ($domains.DomainName -notcontains $domain) {
                $externalRecipients.Add($email) | Out-Null
            }    
        }
 
        if ($externalRecipients) {
            $extRecString = $externalRecipients -join ", "
            Write-Host "$($rule.Name) forwards to $extRecString" -ForegroundColor Yellow
 
            $ruleHash = New-Object System.Collections.ArrayList
            $ruleHash = "" | Select-Object "PrimarySmtpAddress", "DisplayName", "RuleId", "RuleName", "ExternalRecipients"
            $ruleHash.PrimarySmtpAddress = $mailbox.PrimarySmtpAddress
            $ruleHash.DisplayName = $mailbox.DisplayName
            $ruleHash.RuleId = $rule.Identity
            $ruleHash.RuleName = $rule.Name
            $ruleHash.ExternalRecipients = $extRecString
            $RuleDb.Add($ruleHash) | Out-Null
        }
    }
}
<#Body = $RuleDb | ConvertTo-Html -Title "External Mail Forwarding Report" -PostContent "<p>Creation Date: $(Get-Date)<p>" | Out-String#>
$RuleDb | ConvertTo-Html -Title "External Mail Forwarding Report" -PostContent "<p>Creation Date: $(Get-Date)<p>" | Out-File -FilePath c:\temp\ForwardedEmailRuleM365$Rundate.html
$RuleDb | Export-Csv -Path c:\temp\ForwardedEmailRuleM365$Rundate.csv -NoTypeInformation

<#Send-MailMessage -From $SMTPFrom -To $SMTPTo -Subject $Subject -BodyAsHtml $Body -UseSsl -Credential $LiveCred -SmtpServer $SMTPServer -Port $SMTPPort -Attachments c:\temp\ForwardedEmailRuleM365$Rundate.csv #>

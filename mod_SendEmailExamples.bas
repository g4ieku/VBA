Attribute VB_Name = "mod_SendEmailExamples"
Option Explicit


' Module description: Examples of using the SendEmail sub (in module modSendEmailHelper) to configure your email in just one line

'SUBS
'=========
'
'SendEmailExample       Examples of sending emails with multiple attachments
'MultiRecipients        Examples of sending emails with multiple recipients
'MultiAttachments       Examples of sending emails with multiple attachments
'SendMulti              Send multiple emails

' Examples of sending emails with multiple attachments
Sub SendEmailExample()
    
    ' Basic email example
    Call SendEmail("Who@gmail.com", "subject", "body")

    ' Include signature
    Call SendEmail("Who@gmail.com", "subject", "body", includeSignature:=True)
    
    ' Recipients/Attachments from range and signature
    Call SendEmail("Who@gmail.com", "subject", "body", includeSignature:=True _
                    , recipientsRange:=shEmails.Range("A2:A6") _
                    , attachmentsRange:=shEmails.Range("C2:C6"))
End Sub


' Examples of sending emails with multiple recipients
Sub MultiRecipients()
    
    ' Add recipients as comma seperated list
    Call SendEmail("Who@gmail.com,Who2@gmail.com" _
            , "RecipientsString", "body")
         
    ' Add recipients from a range
    Call SendEmail("", "RecipientsRange", "body" _
                        , recipientsRange:=shEmails.Range("A2:A6"))
                        
    Dim arr As Variant
    arr = Array("Who@gmail.com", "Who2@gmail.com")
                   
    ' Add recipients from an array
    Call SendEmail("Who@gmail.com", "RecipientsArray", "body" _
                        , recipientsArray:=arr)
                        
End Sub
    
' Examples of sending emails with multiple attachments
Sub MultiAttachments()
    
    ' Attachments by range
    Call SendEmail("Who@gmail.com", "AttachRange", "body" _
                        , attachmentsRange:=shEmails.Range("C2:C6"))
                        
    ' Attachments by array
    Dim arr As Variant
    arr = Array("c:\temp\attach1.xlsx", "c:\temp\attach2.xlsx")
    Call SendEmail("Who@gmail.com", "AttachArray", "body" _
                        , attachmentsArray:=arr)
                        
    ' Attachments by string
    Call SendEmail("Who@gmail.com", "AttachString", "body" _
                        , attachmentsString:="c:\temp\attach1.xlsx,c:\temp\attach2.xlsx")
End Sub

' Sub description: Send one email per recipient in a range

' Send multiple emails
Sub SendMulti()

    Dim emailAddress As Variant
    For Each emailAddress In shEmails.Range("A2:A3").Value
        
        Call SendEmail(Trim(emailAddress), "Subject", "Body")
                
    Next emailAddress
    
End Sub



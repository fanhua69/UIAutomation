# UIAutomation
UIAutomation exercises

1. Using UIAutomation to start the process of "Outlook"
2. Using UIAutomation to simulate the clicking of "New Email"
3. Using UIAutomation to simulate to put an email address to the "To" edit box.
4. Using UIAutomation to simulate to put an email address to the "Subject" edit box.
5. Using UIAutomation to simulate the clicking of "Send"

Problems could be further investigation:
1. The process of "Outlook" could not be found by the process ID.
2. The email content could not be filled with "Value" pattern. 
   What I can see the difference of the email body and the "Subject" is that 
   the "Subject" is keyboard "Focusable", while the email body is "Unfocusable".
   I tried to "SendMessage" which didn't succeed, also "DOM" of the email body
   window is not available.
   I investigated "TextRange" pattern, but not found any interface to modify the text.
   
   
   
Thanks.

David 

fanhua69@gmail.com

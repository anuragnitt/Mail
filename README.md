# E-Mail
* Python scripts for sending and fetching E-Mails.
* Connects to Gmail account for the same.
* Requires allowed access to "less secure apps". https://myaccount.google.com/lesssecureapps
* Requires "Enable IMAP", "Auto-Expunge off" and "Immediately delete the message forever" options enabled in Gmail settings. https://mail.google.com/mail/u/0/#settings/fwdandpop

# smtp_gmail.py
* Uses SMTP protocol for sending E-Mails. (smtplib library)
* Choice to use any among SSL and TLS network traffic encryptions.
* Option to encrypt (AES-256) the body and attachments of the outgoing E-Mails.
* Unique encryption key stored remotely. (Binary file generated in case export of key fails)
* Password (unique for each E-Mail) generated for receivers to access the encrypted E-Mails.

# imap_gmail.py
* Uses IMAP protocol for fetching E-Mails from Gmail inbox. (imaplib library)
* Uses SSL network traffic encryption.
* Choice to use password or binary file for decryption which are unique for each E-Mail.
* Limited no. of attempts allowed to decrypt the E-Mail. Failure results in deletion of E-Mail.
* Not for "syncing" the offline and online inboxes.

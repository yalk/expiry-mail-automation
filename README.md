expiry-mail-automation
======================

This is a README for MailExpiringContacts.plx file.

Brief description:
==================
	The script is designed to read an excel file and mail the contacts/rows which qualify to a predefined condition.

Detailed description:
=====================
		The script will read the excel file and and look for 3 columns (email,created,duration), when found the script will
	save the row and column(is redundant in some cases) index and then use it later to reference the cells which fall
	under that column.
		The script then starts to check the rows one by one and see if the sum of startDate and duration (that is
	the date at which the service started + the duration for which it was requested) minus the date today is <=7
	((startDate+duration)-today<=&) if the result is true, then the email is validated (with regular expression)
	and then added to an array which stores all the email address.
		The next segment just sends emails to the email addresses stored in the same array. Using outlook. An object is created,
	files are read (which contain the subject and the body of the email) and then the mail is sent.
		Other features include printing of a table on the command prompt which tells the user that these users qualify and
	their service is ending soon. Another list is printed which tells the user these are the email address to which the email was
	finally sent (in case the email was wrong/wasn't accepted by the regex it isn't included in the array).
		Everything which is printed on the screen is also saved in a log to be referenced later in case one needs to check.
	

Following things should be noted/implemented to successfully run the script.
============================================================================

1.	Columns in the excel file should be in the following order:
		email,created,duration

2.	The columns should contain these phrases to be identified correctly
		Email column:					emailid OR mailid OR email
		Date of creation column:		created
		Duration of service:			duration
		
		Note: 	A regular expression is used to validate the columns.
				A sample .xlsx file which works perfectly with the script is provided.

3.	I had troubles installing Mail::Outlook module from cpanm (I was using perl v5.12.3 built for MSWin32-x86-multi-thread)
	A user in #perl on IRC suggested the following steps which solved the problem:
		a.	Download unzip binary(http://kaz.dl.sourceforge.net/project/gnuwin32/unzip/5.51-1/unzip-5.51-1-bin.zip)
		b.	unzip it
		c.	copy \unzip\bin\unzip
		d.	paste in C:\Strawberry\c\bin\

4.	Log files aren't useless, don't delete them, every file included has some significance.
	List of files used for logs and mailing
		LOGS:
			EVERYTIME A PROGRAM IS RUN, LOG IS UPDATED WITH A NEW SEGMENT WHICH BEGINS WITH: 100x=\nDATE & TIME\n100x=
				log_mailsent.txt	:	contains emails of contacts which were sent mail
				log_entry.txt		:	contains list of qualifying entries in the excel sheet which should be sent a mail
		
		Note:	They should be created automatically even if you deleted them.

5. 	Email subject and body can be edited by editing the following files which should be in the same directory.
		EMAIL RELATED:
			mail_subject.txt	:	contains the email subject
			mail_body.txt		:	contains the email body

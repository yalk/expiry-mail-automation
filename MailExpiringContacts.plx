#use warnings;
use strict;

#modules to be used
use Spreadsheet::XLSX;
use DateTime::Format::Excel;
use DateTime;
use Time::Piece;
use Mail::Outlook;
use Path::Class;
use autodie;                  #die if problem reading or writing a file

#
#list of files used for logs and mailing
#LOGS:
#	EVERYTIME A PROGRAM IS RUN, LOG IS UPDATED WITH A NEW SEGMENT WHICH BEGINS WITH: 100x=\nDATE & TIME\n100x=
#		log_mailsent.txt	:	contains emails of contacts which were sent mail
#		log_entry.txt		:	contains list of qualifying entries in the excel sheet which should be sent a mail
#
#EMAIL RELATED:
#		mail_subject.txt	:	contains the email subject
#		mail_body.txt		:	contains the email body
#

#open log files
my $filename1 = 'log_mailsent.txt';
open(my $log_mailsent, '>>', $filename1) or die "Could not open file '$filename1' $!";
my $filename2 = 'log_entry.txt';
open(my $log_entry, '>>', $filename2) or die "Could not open file '$filename2' $!";

#paste date and time in logs with header design
my $pasteDateInLog = localtime->strftime();

#paste in log_entry.txt
say $log_entry "\n","="x100;
say $log_entry "$pasteDateInLog";
say $log_entry "="x100;

#paste in log_mailsent.txt
say $log_mailsent "\n","="x100;
say $log_mailsent "$pasteDateInLog";
say $log_mailsent "="x100;

#open email files and copy content
my $dir = dir("./");																#./ =current directory
my $file = $dir->file("mail_subject.txt");					#define the file name
my $emailSubject = $file->slurp();									#Read in the entire contents of a file

my $dir = dir("./");
my $file = $dir->file("mail_body.txt");
my $emailBody = $file->slurp();

my $excel = Spreadsheet::XLSX -> new ('test.xlsx', my $converter);
                                                    #replace test.xlsx by the appropriate file name
my $emailRow=0;																			#row index of email cell
my $emailCol=0;																			#column index of email cell
my $createdRow=0;																		#row index of date of creation
my $createdCol=0;																		#column index of date of creation
my $durationRow=0;																	#row index of duration (in months)
my $durationCol=0;																	#column index of duration (in months)
my @listOfEmailIDs;																	#array of email IDs to be sent a mail

OUTERMOST:																				  #label to help in quitting loop(s) when required
for my $sheet (@{$excel -> {Worksheet}})
{
	#capture column and row indexes of email, duration and creation date cells
	MOVETONEXT:																			  #label to help in quitting loop(s) when required
    {
		$sheet -> {MaxRow} ||= $sheet -> {MinRow};
        foreach my $row ($sheet -> {MinRow} .. $sheet -> {MaxRow})
                                                    #loop starting from min row to max row
        {
			$sheet -> {MaxCol} ||= $sheet -> {MinCol};
            foreach my $col ($sheet -> {MinCol} ..  $sheet -> {MaxCol})
                                                    #loop starting from min col to max col
            {
				my $cell = $sheet -> {Cells} [$row] [$col];
                if ($cell)
                {
                    my $regexCell =$cell -> {Val};
                    $regexCell =~ s/[\s-]//g;					#regular expression aka regex to remove whitespace and hyphon '-'
                    if ($regexCell =~ /emailid|mailid|email/gi)					
                                                      #regex to find email cell
					{
						#storing the row and col index
						$emailRow=$row;													
						$emailCol=$col;
					}
					if ($regexCell =~ /created/gi)							#regex to find creation date
					{
						#storing the row and col index
						$createdRow=$row;
						$createdCol=$col;
					}
					if ($regexCell =~ /duration/gi)							#regex to find duration of service
					{
						#storing the row and col index
						$durationRow=$row;
						$durationCol=$col;
						last MOVETONEXT;												  #break loop and move on to next step
					}
				}
			}
		}
	}
	
	my $ctr=0;
	print"\nEntries whose service time has expired:\n\nStart\t\tEnd\t\tToday\t\tDiff\tEmail\n";
																						          #table heading in cmd
	say $log_entry "Start\t\tEnd\t\tToday\t\tDiff\tEmail\n";
	                                                    #paste table heading log file
	$sheet -> {MaxRow} ||= $sheet -> {MinRow};
	foreach my $row (1 .. $sheet -> {MaxRow})
	{
		my $cell = $sheet -> {Cells} [$row] [$createdCol];
		my $duration= $sheet -> {Cells} [$row] [$durationCol];
		my $emailID= $sheet -> {Cells} [$row] [$emailCol];
		
		if ($cell)
		{
			#convert value in cell to string and store in variable
			my $serviceStartDay = $cell -> {Val};
			
			#convert excel format date to human readable format
			my $excel = DateTime::Format::Excel->new();
			my $startDate= $excel->parse_datetime($serviceStartDay)->dmy;
			                                                #human readable date at which service started
			
			#calculate when will service end, convert given date in $startDate + $duration of subscription (from excel) to excel format date
			my $endDate = DateTime->new(day => (substr $startDate, 0, 2),month=> ((substr $startDate, 3, 2)+$duration->{Val}-1),year=>(substr $startDate, 6, 4));
			my $serviceEndDay = DateTime::Format::Excel->format_datetime( $endDate );	
			
						
			#convert today's date to excel format date
			my $todayDate = DateTime->new( year => localtime->strftime('%Y'), month => localtime->strftime('%m'), day => localtime->strftime('%d') );
			my $todayDay = DateTime::Format::Excel->format_datetime( $todayDate );		
			
			my $diff=$serviceEndDay-$todayDay;							#calculate difference between end day and today

			if($diff<=7)																    #if number of days for service to get over is less than 7 then add $emailID to array and print
			{
				my $tempEmail= $emailID->{Val};
				print ((substr $startDate, 0, 10),"\t",(substr $endDate->dmy, 0,10),"\t",(substr $todayDate->dmy, 0, 10),"\t$diff\t",$tempEmail,"\n");
																						          #print rows of the table
				say $log_entry ((substr $startDate, 0, 10),"\t",(substr $endDate->dmy, 0,10),"\t",(substr $todayDate->dmy, 0, 10),"\t$diff\t",$tempEmail); 
																						          #add things printed on screen to the logs
				if ($emailID->{Val} =~ /[A-Z]+\.[A-Z]+\@yourcompany+\.domain/i)
				                                              #regex to validate email: string1.string2@yourcompany.domain 'i' in the end ignores case
																						          #can be replaced with /[A-Z0-9._%+-]+@[A-Z0-9.-]+\.[A-Z]{2,4}/i for any email
				{
					push @listOfEmailIDs,$emailID->{Val};				#push email address in the array
				}
				else
				{
					print "Invalid $emailID at row ", $row+1," in excel sheet\n";	
					                                            #prompt cell location in excel sheet to the user error in email id
				}
			}
		}
	}
	close $log_entry;																	  #close file
	last OUTERMOST;																		  #kill the loop
}

if (@listOfEmailIDs != 0)															#proceed into if block only when the array is not empty
{
	print "Send email to above mentioned contacts?(yes/no): ";
	                                                    #validate step from the user
	my $sendEmailInput=<STDIN>;															
	if($sendEmailInput=~ /^yes$/i)											#if input matches exactly with yes (i for case insensitive)
	{
		my $outlook = new Mail::Outlook();												
		print "\nSuccessfully sent email to following contacts:\n"; 					
		foreach my $emailid(@listOfEmailIDs)											
		{
			my $message = $outlook->create();											
			$message->To($emailid);														
			#$message->Cc('Them <them@example.com>');				#edit this if required
			#$message->Bcc('Us <us@example.com>; anybody@example.com');
			                                                #edit this if required
			$message->Subject($emailSubject);								#paste variable content in member of object
			$message->Body($emailBody);											#paste variable content in member of object
			#$message->Attach(@lots_of_files);							#edit this if required
			#$message->Attach(@more_files);    							#attachments are appended
			#$message->Attach($one_file);      											#so multiple calls are allowed
			$message->send;																
			print "$emailid\n";															
			say $log_mailsent "$emailid\n";												
		}
	}
	else																				        #if input doesn't match with yes then enter this block
	{
		print "Email not sent\n";														
	}
	close $log_mailsent;																#close file
}

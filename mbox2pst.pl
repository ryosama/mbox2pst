#!/usr/bin/perl
use Email::PST::Win32;
use Data::Uniqid qw(luniqid); 	# copy message into a temp file
use File::Find;
use File::Path;
use strict;
use File::Basename;
use Cwd;
use Getopt::Long;				# manage options
$|=1;

my ($mboxdir,$pst,$tempdir,$help,$quiet);
GetOptions(	'mboxdir=s'=>\$mboxdir, 'pst=s'=>\$pst, 'tempdir=s'=>\$tempdir, 'help|usage!'=>\$help, 'quiet!'=>\$quiet) ;

my $program_dir = getcwd;
print "Working directory : $program_dir\n" unless $quiet;

die <<EOT if ($help);
Options :
--mboxdir=<mbox_input_directory> (required)
	The mbox directory you want to convert

--pst=<output_pst_file>
	The PST file converted (default is $program_dir/out/Outlook.pst)

--tempdir=<temporary_directory>
	The temp directory (default is $program_dir/tmp)

--usage or --help
	Display this message

--quiet
	Don't print anything
EOT

die "Option --mboxdir=<mbox_input_directory> is needed" unless length($mboxdir)>0;
die "$mboxdir is not a directory" if !-d $mboxdir;
$mboxdir =~ s|\\|\/|g; # convert \ to /
$mboxdir =~ s|/$||; # remove last /

print "Input directory : $mboxdir\n" unless $quiet;

$tempdir = $tempdir || "$program_dir/tmp";
mkpath($tempdir);

$pst = $pst || "$program_dir/out/Outlook.pst";
mkpath(dirname($pst));
print "Export PST file : $pst\n" unless $quiet;

# clean up last run
print "Clean up last run... " unless $quiet;
unlink($pst);
find(\&clean_last_run, $tempdir);
sub clean_last_run {
	return if ($_ eq '.' || $_ eq '..');
    unlink($_) if (/\.mail$/i);
}
print "OK" unless $quiet;

# creat new PST file
my $pst = Email::PST::Win32->new(
    filename     => $pst,
    display_name => 'Export from mbox2pst',
);

# Errors may occur when high numbers of items are added.
# A count_per_session > 0 will determine when to close and
# reopen the PST file. The default value is 1000.
$pst->count_per_session( 2000 );
 
# Get number of MIME files added
my $count = $pst->instance_counter;

my $buffer = ''; # buffer for current message
my $total_message = 0;

my $mbox = '';
my $outlook_box = '';

find(\&extractMboxToFiles, $mboxdir);
sub extractMboxToFiles {
    $mbox = $File::Find::name;
    $mbox =~ s/^$mboxdir\/?//; # remove leading path

	return if ($_ eq '.' || $_ eq '..' || -d $_ || /\.msf$/ || $_ eq 'filterlog.html' || $_ eq 'msgFilterRules.dat');

    my $total_message = 0;

    print "\nExtracting message from '$mbox' " unless $quiet;

    # Add an MIME file to the PST
    open(F,"<$_") or die "Could not open '$mbox' ($!)";
    while(my $line = <F>) { # foreach line
        if ($line =~ /^From\s+-\s+\w{3}\s+(\w{3})\s+(\d{2})\s+\d{2}:\d{2}:\d{2}\s+(\d{4})\b/i) {  # new message syntax : From - Mon Jan 05 08:37:43 2012
            if ($total_message>0) {
                print "$total_message..." if ($total_message % 50 == 0 && !$quiet);
                my $uniqid 	= luniqid;
                open(OUTPUT, "+>$tempdir/$uniqid.mail");
                print OUTPUT $buffer; # write buffer into the output file
                close OUTPUT;
                $buffer = '';		# reset buffer
            }
            $total_message++;
        } # end if new message

        $buffer .= $line ; # put line into buffer
    } # foreach line

    $outlook_box = $mbox ;
    $outlook_box =~ s/\.sbd//g;
    $outlook_box =~ s/Inbox/Boîte de réception/;
    $outlook_box =~ s/Sent/Élements envoyés/;
    $outlook_box =~ s/Trash/Élements supprimés/;
    $outlook_box =~ s/Junk/Courrier indésirable/;
    $outlook_box =~ s/Draft/Brouillons/;
    print "\nAdd message from '$mbox' to '$outlook_box'\n" unless $quiet;

    find(\&add_to_pst, $tempdir);
}

# add *.mail to pst file
sub add_to_pst {
	return if ($_ eq '.' || $_ eq '..');
    if (/\.mail$/i) {
        $pst->add_mime_file( $_, $outlook_box, 1 ? 'note' : 'post' );
        unlink($_);
    }
}

# Close the PST file
$pst->close;
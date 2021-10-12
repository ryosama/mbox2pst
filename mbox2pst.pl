#!/usr/bin/perl
use strict;
use Email::PST::Win32;
use Data::Uniqid qw(luniqid); 	# copy message into a temp file
use File::Find;
use File::Path;
use File::Slurp;
use File::Basename;
use Cwd;
use Getopt::Long;				# manage options
use Data::Dumper;

# for gui
use Win32::GUI();
use Win32::GUI::Constants qw(CW_USEDEFAULT WS_OVERLAPPEDWINDOW WS_DISABLED WS_POPUP WS_CAPTION WS_THICKFRAME WS_EX_TOPMOST SW_HIDE);
use Win32::GUI::BitmapInline;

$|=1;

our ($mboxdir,$pst,$tempdir,$gui,$help,$quiet);
GetOptions(	'mboxdir=s'=>\$mboxdir, 'pst=s'=>\$pst, 'tempdir=s'=>\$tempdir, 'gui!'=>\$gui,'help|usage!'=>\$help, 'quiet!'=>\$quiet) ;

our ($main,$mainParameters,%assets,$program_dir,$outlook_box,$pstObj,$total_message,$current_message,$mbox,$total_mbox_files) ;
$total_mbox_files = 0;
our $y=10;


# create gui if needed
if ($gui) {
    # hide dos box
    my $DOS = Win32::GUI::GetPerlWindow();    
    Win32::GUI::Hide($DOS);

    # get thunderbird mail directory
    find(\&getPrefFile, $ENV{'APPDATA'}.'\\Thunderbird\\Profiles\\');

    require('./gui.pl'); # draw GUI
} else {
    do_convert();
}


# get pref file for thunderbird
sub getPrefFile {
    return if ($_ ne 'prefs.js');
    my @lines = read_file($_);
    foreach my $l (@lines) {
        # user_pref("mail.server.server1.directory", "path to mail folder");
        if ($l =~ /^\s*user_pref\(\s*['"]mail.server.server1.directory["']\s*,\s*['"](.+?)['"]\s*\)\s*;/) {
            $mboxdir = $1;
            $mboxdir =~ s|\\+|\\|g; # convert \\ to \
            return; # break
        }
    }
}


# main part
sub do_convert {
    $main->StatusBar->Text("Convertion en cours") if $gui;

    $program_dir = getcwd;
    print "Working directory : $program_dir\n" unless $quiet;

    display_help() if $help;

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
    print "OK" unless $quiet;

    # creat new PST file
    $pstObj = Email::PST::Win32->new(
        filename     => $pst,
        display_name => 'Export from mbox2pst',
    );

    # Errors may occur when high numbers of items are added.
    # A count_per_session > 0 will determine when to close and
    # reopen the PST file. The default value is 1000.
    $pstObj->count_per_session( 2000 );
    
    # Get number of MIME files added
    my $count = $pstObj->instance_counter;

    $total_message = 0; # total number of message in current mbox file
    $outlook_box = ''; # name of outlook box (inbox, trash, ...)
    $current_message = 0; # index of current message

    # count number of mbox files
    if ($gui) {
        find(\&countMboxs, $mboxdir);
        $main->ProgressBarTotal->SetRange(0,$total_mbox_files);
        print "total_mbox_files=$total_mbox_files\n";
    }

    find(\&extractMboxToFiles, $mboxdir);

    # fini
    $main->StatusBar->Text("Convertion termin\x{e9}e") if $gui;

    # Close the PST file
    $pstObj->close;
} # end of do_convert


# display command line options
sub display_help {
    die <<EOT ;
Options :
--gui
    Launch user interface

--mboxdir=<mbox_input_directory> (required if no GUI)
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
}


# remove *.mail and pst files from last run
sub clean_last_run {
	return if ($_ eq '.' || $_ eq '..');
    unlink($_) if (/\.mail$/i);
}


# count the number of files to convert (for progress bar)
sub countMboxs {
	return if ($_ eq '.' || $_ eq '..' || -d $_ || /\.msf$/ || $_ eq 'filterlog.html' || $_ eq 'msgFilterRules.dat');
    $total_mbox_files++;
}


# extract an mbox file to *.mail files
sub extractMboxToFiles {
    my $buffer = '';
    $mbox = $File::Find::name;
    $mbox =~ s/^$mboxdir\/?//; # remove leading path

	return if ($_ eq '.' || $_ eq '..' || -d $_ || /\.msf$/ || $_ eq 'filterlog.html' || $_ eq 'msgFilterRules.dat');

    $main->StatusBar->Text("$mbox : Extraction des mails") if $gui;

    $total_message = 0;

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

    # get convert name rules
    my @name_convert_rules = split /[\r\n]+/, $mainParameters->TextfieldParameters->Text();
    foreach my $rule (@name_convert_rules) {
        $rule =~ s/^\s*//; # ltrim
        $rule =~ s/\s*$//; # rtrim
        my ($from,$to) = split(/\s*=>\s*/, $rule);
        $outlook_box =~ s/$from/$to/;
    }

    print "\nAdd message from '$mbox' to '$outlook_box'\n" unless $quiet;

    if ($gui) {
        $main->ProgressBarPartial->SetRange(0,$total_message-1);
        $main->ProgressBarPartial->SetPos(0);
    }

    $current_message = 1;
    find(\&add_to_pst, $tempdir);

    $main->ProgressBarTotal->StepIt() if $gui; # update global progress bar
}


# add *.mail to pst file
sub add_to_pst {
	return if ($_ eq '.' || $_ eq '..');
    if (/\.mail$/i) {
        $pstObj->add_mime_file( $_, $outlook_box, 1 ? 'note' : 'post' );
        unlink($_);
        $current_message++;

        if ($gui) { # update partial progress bar
            $main->ProgressBarPartial->StepIt();
            $main->StatusBar->Text("$mbox : message $current_message / $total_message ".sprintf('(%0.1f %%)',($current_message / $total_message) * 100));
        }
    }
}


############################################################# EVENT GUI ##########################################

# select directory to convert
sub ButtonChooseInputDir_Click {
    my $dir = Win32::GUI::BrowseForFolder (
        -title     => "Choissiez un répertoire mbox",
        -directory => $mboxdir || $ENV{'HOME'}.'\\.Mail',
        -folderonly => 1,
    );
    $main->TextfieldInputDir->Text($dir);
}


# select output PSF file
sub ButtonChooseOutputFilename_Click {
	my @file = Win32::GUI::GetOpenFileName(
		-filter => ['PST - Outlook format', '*.pst',
					'All files - *.*', '*'
					],
		-directory => $ENV{'HOME'}.'\\Desktop',
		-title => 'Choissiez un fichier de sortie PST',
		-file => 'Outlook.pst'
	);
	$main->TextfieldOutputFilename->Text($file[0]);
}

# do convert
sub ButtonConvert_Click {
    $mboxdir = $main->TextfieldInputDir->Text();
    $pst     = $main->TextfieldOutputFilename->Text();
	do_convert();   
}
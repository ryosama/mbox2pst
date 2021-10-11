#!/usr/bin/perl
use Email::PST::Win32;
use Data::Dumper;
use Data::Uniqid qw(luniqid); 	# copy message into a temp file
use File::Find;
use File::Path;
use strict;
use File::Basename;
use Cwd;
use Getopt::Long;				# manage options

# for gui
#use Win32::API;
use Win32::GUI();
use Win32::GUI::Constants qw(CW_USEDEFAULT WS_OVERLAPPEDWINDOW WS_DISABLED WS_POPUP WS_CAPTION WS_THICKFRAME WS_EX_TOPMOST SW_HIDE);
use Win32::GUI::BitmapInline;

$|=1;

my ($mboxdir,$pst,$tempdir,$gui,$help,$quiet);
GetOptions(	'mboxdir=s'=>\$mboxdir, 'pst=s'=>\$pst, 'tempdir=s'=>\$tempdir, 'gui!'=>\$gui,'help|usage!'=>\$help, 'quiet!'=>\$quiet) ;

our ($main,%assets,$program_dir,$outlook_box,$pstObj,$total_message,$current_message,$mbox) ;
our $y=10;

# cree la ffenetre principale
if ($gui) {
    my $DOS = Win32::GUI::GetPerlWindow();    
    Win32::GUI::Hide($DOS);
    require('./gui.pl');
} else {
    do_convert();
}


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

    find(\&extractMboxToFiles, $mboxdir);

    # fini
    $main->StatusBar->Text("$mbox : Convertion termin\x{e9}e") if $gui;

    # Close the PST file
    $pstObj->close;
} # end of do_convert

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

sub clean_last_run {
	return if ($_ eq '.' || $_ eq '..');
    unlink($_) if (/\.mail$/i);
}


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
    $outlook_box =~ s/Inbox/Boîte de réception/;
    $outlook_box =~ s/Sent/Élements envoyés/;
    $outlook_box =~ s/Trash/Élements supprimés/;
    $outlook_box =~ s/Junk/Courrier indésirable/;
    $outlook_box =~ s/Draft/Brouillons/;
    print "\nAdd message from '$mbox' to '$outlook_box'\n" unless $quiet;

    $current_message = 1;
    $main->ProgressBar->SetRange(0,$total_message) if $gui;
    find(\&add_to_pst, $tempdir);
}


# add *.mail to pst file
sub add_to_pst {
	return if ($_ eq '.' || $_ eq '..');
    if (/\.mail$/i) {
        $pstObj->add_mime_file( $_, $outlook_box, 1 ? 'note' : 'post' );
        unlink($_);
        $current_message++;
        updateProgressBar() if $gui;
    }
}


############################################################# EVENT GUI ##########################################

# selectionne le répertoire a convertir
sub ButtonChooseInputDir_Click {
    my $dir = Win32::GUI::BrowseForFolder (
        -title     => "Choissiez un répertoire mbox",
        -directory => $ENV{'HOME'}.'\\.Mail',
        -folderonly => 1,
    );
    $main->TextfieldInputDir->Text($dir);
}


# selectionne le fichier de sortie CSV
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

# selectionne le fichier de sortie CSV
sub ButtonConvert_Click {
    $mboxdir = $main->TextfieldInputDir->Text();
    $pst     = $main->TextfieldOutputFilename->Text();
	do_convert();   
}

sub updateProgressBar {
    #my $pourcentage = ($current_message / $total_message) * 100;
    Win32::GUI::DoEvents() >= 0;
    $main->ProgressBar->SetPos($current_message);
    $main->StatusBar->Text("$mbox : message $current_message / $total_message ".sprintf('(%0.1f %%)',($current_message / $total_message) * 100));
    #$main->ProgressBar->StepIt();
    #print "update progressbar to \$current_message=$current_message / \$total_message=$total_message) pourcentage=$pourcentage\n";
}
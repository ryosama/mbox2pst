############################################ Load assets (icons and bitmaps) ############################################
foreach (qw/folder_16x16.ico file_16x16.ico gear_32x32.ico/)  {
    my $path = '';
    if (-e "gfx/$_") {
        $path = "gfx/$_";
    } elsif (-e "$ENV{PAR_TEMP}/inc/gfx/$_") {
        $path = "$ENV{PAR_TEMP}/inc/gfx/$_";
    }
    # load ico or bmp
    if (/\.ico$/) {
        $assets{$_} = Win32::GUI::Icon->new( $path );
    } elsif (/\.bmp$/) {
        $assets{$_} = Win32::GUI::Bitmap->new( $path );
    }
}

# main window
$main = Win32::GUI::Window->new(
    -name => 'Main',
    -text => 'mbox2pst',
    -width => 520,
    -height=> 280,
    # -menu   => $menu,
    -left => CW_USEDEFAULT
);
$main->Center();

# status bar for messages
$main->AddStatusBar(
    -name => 'StatusBar',
    -text => '',
    -width => 600,
    -height => 20
);

# label for input directory
$main->AddLabel(
    -pos => [ 20, $y ],
    -text => "Répertoire mbox"
);


# text field for input directory
$main->AddTextfield(
    -name 	=> 'TextfieldInputDir',
    -pos 	=> [100, $y],
    -size	=> [300, 20],
    -text 	=> $mboxdir || $ENV{'HOME'}.'\\.Mail'
);


# button to chosse input directory
$main->AddButton(
    -name 	=> 'ButtonChooseInputDir',
    -pos    => [ 410, $y-5 ],
    -size   => [28,28],
    -icon   => $assets{'folder_16x16.ico'}
);

# parameters button
$main->AddButton(
    -name 	=> 'ButtonParameters',
    -pos    => [ 450, $y-5 ],
    -size   => [45,40],
    -icon   => $assets{'gear_32x32.ico'},
    -onClick=> sub {
        $mainParameters->Top( $main->Top() );
		$mainParameters->Left( $main->Left() + $main->Width() );
		#$mainParameters->Height( $main->Height() );
		$mainParameters->Show();
    }
);

$y+=30;


# label for output file
$main->AddLabel(
    -pos => [ 20, $y ],
    -text => 'Fichier PST'
);

# text field for output file
$main->AddTextfield(
    -name 	=> 'TextfieldOutputFilename',
    -pos 	=> [100, $y],
    -size	=> [300, 20],
    -text 	=> $ENV{'HOME'}.'\\Desktop\\Outlook.pst'
);

# button to choose output file
$main->AddButton(
    -name 	=> 'ButtonChooseOutputFilename',
    -pos    => [ 410, $y-5 ],
    -size   => [28,28],
    -icon   => $assets{'file_16x16.ico'}
);

$y+=30;

$main->AddCheckbox(
    -name => 'CheckboxExcludeTrash',
    -text => 'Exclure la poubelle',
    -checked=> 1,
    -pos => [ 20 , $y ]
);

$main->AddCheckbox(
    -name => 'CheckboxExcludeJunk',
    -text => 'Exclure les indésirables',
    -checked=> 1,
    -pos => [ 160 , $y ]
);

$y+=30;

# main button for convert
$main->AddButton(
	-name => 'ButtonConvert',
	-pos => [ 20, $y ],
	-text => "Convertir",
    -size => [460, 50],
);

$y+=60;

# progress bar 1 (partial)
$main->AddProgressBar(
	-name => 'ProgressBarPartial',
	-pos => [ 20, $y ],
    -background=>[0,255,85],
    -smooth   => 1,
    -size => [460, 20],
);
$main->ProgressBarPartial->SetStep(1); # increase 1 by 1

$y+=30;

# progress bar 2 (global)
$main->AddProgressBar(
	-name => 'ProgressBarTotal',
	-pos => [ 20, $y ],
    -background=>[0,255,85],
    -smooth   => 1,
    -size => [460, 20],
);
$main->ProgressBarTotal->SetStep(1); # increase 1 by 1
$main->ProgressBarTotal->SetPos(0); # increase 1 by 1




# parameters window
$mainParameters = Win32::GUI::Window->new(
	-name  => "mainParameters",
	-title => "Paramètres",
	-pos   => [ 0,0 ],
	-size  => [ 300, 270 ],
);

# label for output file
$mainParameters->AddLabel(
    -pos => [ 0, 10 ],
    -text => 'Convertion des noms de répertoire (regex possible)'
);


# textarea pour la sortie du programme externe
$mainParameters->AddTextfield(
	-name 		=> 'TextfieldParameters',
	-pos 		=> [ 0, 30 ],
    -size	    => [300, 200],
	-multiline 	=> 1,
	-hscroll   	=> 1,
	-vscroll   	=> 1,
	#-autohscroll=> 1,
	#-autovscroll=> 1,
	-tabstop 	=> 1,
	-readonly   => 0,
    -text =>    $rules
);


# display main window
$main->Show();
Win32::GUI::DoEvents();
Win32::GUI::Dialog();

exit(0);


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
    push @exclude, 'Trash' if $main->CheckboxExcludeTrash->Checked();
    push @exclude, 'Junk'  if $main->CheckboxExcludeJunk->Checked();
	do_convert();   
}
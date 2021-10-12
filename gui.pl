############################################ Load assets (icons and bitmaps) ############################################
foreach (qw/folder_16x16.ico file_16x16.ico/)  {
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


$main = Win32::GUI::Window->new(
    -name => 'Main',
    -text => 'mbox2pst',
    -width => 520,
    -height=> 260,
    # -menu   => $menu,
    -left => CW_USEDEFAULT
);
$main->Center();

# status bar pour les messages
$main->AddStatusBar(
    -name => 'StatusBar',
    -text => '',
    -width => 600,
    -height => 20
);

# choix du répertoire d'input
# libelle pour nom du fichier output
$main->AddLabel(
    -pos => [ 20, $y ],
    -text => "Répertoire mbox"
);


# texte qui contient le fichier de sortie
$main->AddTextfield(
    -name 	=> 'TextfieldInputDir',
    -pos 	=> [100, $y],
    -size	=> [300, 20],
    -text 	=> $mboxdir || $ENV{'HOME'}.'\\.Mail'
);


# bouton pour choisir le fichier de sortie
$main->AddButton(
    -name 	=> 'ButtonChooseInputDir',
    -pos    => [ 410, $y-5 ],
    -size   => [28,28],
    -icon   => $assets{'folder_16x16.ico'}
);

$y+=30;


# libelle pour nom du fichier output
$main->AddLabel(
    -pos => [ 20, $y ],
    -text => 'Fichier PST'
);

# texte qui contient le fichier de sortie
$main->AddTextfield(
    -name 	=> 'TextfieldOutputFilename',
    -pos 	=> [100, $y],
    -size	=> [300, 20],
    -text 	=> $ENV{'HOME'}.'\\Desktop\\Outlook.pst'
);

# bouton pour choisir le fichier de sortie
$main->AddButton(
    -name 	=> 'ButtonChooseOutputFilename',
    -pos    => [ 410, $y-5 ],
    -size   => [28,28],
    -icon   => $assets{'file_16x16.ico'}
);

$y+=30;

# executer le programme externe
$main->AddButton(
	-name => 'ButtonConvert',
	-pos => [ 20, $y ],
	-text => "Convertir",
    -size => [460, 50],
);

$y+=60;

# executer le programme externe
$main->AddProgressBar(
	-name => 'ProgressBarPartial',
	-pos => [ 20, $y ],
    -background=>[0,255,85],
    -smooth   => 1,
    -size => [460, 20],
);
$main->ProgressBarPartial->SetStep(1); # increase 1 by 1

$y+=30;

# executer le programme externe
$main->AddProgressBar(
	-name => 'ProgressBarTotal',
	-pos => [ 20, $y ],
    -background=>[0,255,85],
    -smooth   => 1,
    -size => [460, 20],
);
$main->ProgressBarTotal->SetStep(1); # increase 1 by 1
$main->ProgressBarTotal->SetPos(0); # increase 1 by 1


# affiche la fenetre principale
$main->Show();
Win32::GUI::DoEvents();
Win32::GUI::Dialog();

exit(0);
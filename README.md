# mbox2pst
--------------
Convert mail from mbox format (Thunderbird) to PST (Outlook)

# Redemption.zip
--------------
You should install Redemption DLLs first (unzip and click on install.exe)

# Usage
--------------
Options :

--gui
    Launch user interface
	
--mboxdir=<mbox_input_directory> (required)
	The mbox directory you want to convert

--pst=<output_pst_file>
	The PST file converted (default is program_directory/out/Outlook.pst)

--exclude=<mbox_folder_to_exlude>   (could be repeat)
    To exclude some folders (Junk or Trash for example)

--tempdir=<temporary_directory>
	The temp directory (default is program_directory/tmp)

--usage or --help
	Display this message

--quiet
	Don't print anything

# Examples
------------
`perl mbox2pst.pl --gui`

`perl mbox2pst.pl --mboxdir=c:/users/thunderbird/Mail --pst=Outlook.pst --exclude=Trash --exclude=Junk`


# Parameters (in GUI mode only)
------------
You can change input mbox folder name to another output PST folder name. You can use Regex to specify the input folder.
Example :
- Inbox => Courrier entrant
- Sent  => Email envoyés
- Trash|Deleted  => Poubelle

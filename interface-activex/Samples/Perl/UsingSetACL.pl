#
#	Sample script that demonstrates how to use the SetACL OCX version from Perl.
#
#	Author:			Helge Klein
#
#	Tested with:	ActivePerl 806 on
#						Windows XP SP1
#


#
#	Uses
#
use strict;

use Win32::OLE qw(in with EVENTS);
use Win32::OLE::Const;


#
#	Start of script
#

# 	Create the SetACL object
my $oSetACL				=	Win32::OLE->CreateObject ('SetACL.SetACLCtrl.1') or die "SetACL.ocx is not registered on your system! Use regsvr32.exe to register the control.\n";

#	Load the constants from the type library
my $const				=	Win32::OLE::Const->Load ($oSetACL) or die "Constants could not be loaded from the SetACL type library!\n";

#	Enable event processing
Win32::OLE->WithEvents ($oSetACL, \&EventHandler);

#
#	Check arguments
#
my $Path					=	$ARGV[0] or die "Please specify the (file system) path to use as first parameter!\n";
my $Trustee				=	$ARGV[1] or die "Please specify the trustee (user/group) to use as second parameter!\n";
my $Permission			=	$ARGV[2] or die "Please specify the permission to set as third parameter!\n";

if (! -d $Path)
{
	die "The path specified (1. parameter) does not exist!\n";
}

#	Set the object
my $RetCode				=	$oSetACL->SetObject ($Path, $const->{SE_FILE_OBJECT});

if ($RetCode != $const->{RTN_OK})
{
	die "SetObject failed: " . $oSetACL->GetResourceString ($RetCode) . "\nOS error: " . $oSetACL->GetLastAPIErrorMessage ();
}

#	Set the action
$RetCode					=	$oSetACL->SetAction ($const->{ACTN_ADDACE});

if ($RetCode != $const->{RTN_OK})
{
	die "SetAction failed: " . $oSetACL->GetResourceString ($RetCode) . "\nOS error: " . $oSetACL->GetLastAPIErrorMessage ();
}

#	Add an ACE
$RetCode					=	$oSetACL->AddACE ($Trustee, 0, $Permission, $const->{INHPARNOCHANGE}, 0, $const->{GRANT_ACCESS}, $const->{ACL_DACL});

if ($RetCode != $const->{RTN_OK})
{
	die "AddACE failed: " . $oSetACL->GetResourceString ($RetCode) . "\nOS error: " . $oSetACL->GetLastAPIErrorMessage ();
}

#	Run SetACL
$RetCode					=	$oSetACL->Run ();

if ($RetCode != $const->{RTN_OK})
{
	die "Run failed: " . $oSetACL->GetResourceString ($RetCode) . "\nOS error: " . $oSetACL->GetLastAPIErrorMessage ();
}


#
#	End of script
#

exit 0;

#
#	Functions
#

#	Handle events
sub EventHandler
{
	my ($Object, $Event, @Args)	=	@_;
	
	print "$Args[0]\n";
}
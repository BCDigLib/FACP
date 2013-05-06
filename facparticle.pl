#!C:/Perl/bin/perl -w
use strict;
use IO::File;
use File::Basename qw(basename);
use Win32::OLE qw(in with);
use Win32::OLE::Const 'Microsoft Excel';
use utf8;
use Cwd;

my $filename = $ARGV[0];


$Win32::OLE::Warn = 3; # Die on Errors.

# ::Warn = 2; throws the errors, but #
# expects that the programmer deals  #

my $excelfile=$filename;

my $dir = getcwd;
$dir=~s/\//\\/g;
print "dir is $dir\n";
$excelfile=$dir."//".$excelfile;


print "$excelfile\n";

#First, we need an excel object to work with, so if there isn't an open one, we create a new one, and we define how the object is going to exit

my $Excel = Win32::OLE->GetActiveObject('Excel.Application')
        || Win32::OLE->new('Excel.Application', 'Quit');

#For the sake of this program, we'll turn off all the alert boxes, such as the SaveAs response "This file already exists", etc. using the DisplayAlerts property.

$Excel->{DisplayAlerts}=0;   

#open an existing file to work with 
                                                 
my $Book = $Excel->Workbooks->Open($excelfile);   

#Create a reference to a worksheet object and activate the sheet to give it focus so that actions taken on the workbook or application objects occur on this sheet unless otherwise specified.

my $Sheet = $Book->Worksheets("Sheet1");
$Sheet->Activate();  
#Determine basename of output file to be written in current directory
#
$filename =~ s/xlsx|xls/xml/;
my $output_file = $filename;


#Open the output file; print xml declaration and root node
#
my $fh = IO::File->new($output_file, 'w')
	or die "unable to open output file for writing: $!";
binmode($fh, ':utf8');
$fh->print("<?xml version='1.0' encoding='UTF-8' ?>\n");
$fh->print("<mods:modsCollection xmlns:xlink=\"http://www.w3.org/1999/xlink\" xmlns:xsi=\"http://www.w3.org/2001/XMLSchema-instance\" xmlns:mods=\"http://www.loc.gov/mods/v3\" xsi:schemaLocation=\"http://www.loc.gov/mods/v3 http://www.loc.gov/standards/mods/v3/mods-3-3.xsd\">\n");

##read and process rows

my $usedRange = $Sheet->UsedRange()->{Value};


my $LastRow = $Sheet->UsedRange->Find({What=>"*",
    SearchDirection=>xlPrevious,
    SearchOrder=>xlByRows})->{Row};

print "last row is $LastRow\n";


#my $nextRowID = $Sheet->Range('A'.$LastRow)->{Value};

shift(@$usedRange);

my $CurrentRow=2;

#foreach my $row (@$usedRange){

while (my $row=shift @$usedRange){



#Read a tab-delimited line of metadata and assign each element to an appropriately named variable
#
	my ($wfID, $marcRelatorCode, $authorOrder, $family, $given, $given2, $shortname, $dept, $school, $title, $subtitle, $journalTitle, $enum1, $enum2, $chron2, $chron1, $startPage, $endPage, $pageList, $issn, $type, $url ,$doi, $setText, $ready, $version) = @$row;


	$fh->print("<mods:mods>\n");

### 1. MODS TitleInfo Element

$fh->print("<mods:titleInfo>\n");

##Deal with initial articles
	my $nonsort;
	if ($title =~ m/^The (.*)/) 
		{$nonsort = "The"; 
		$title=$1} 
	elsif ($title =~ m/^A (.*)/) 
		{$nonsort = "A";
		$title=$1} 
	elsif ($title =~ m /^An (.*)/) 
		{$nonsort = "An";
		$title=$1}; 
	
	if ($nonsort) {$fh->print ("\t<mods:nonSort>$nonsort <\/mods:nonSort>\n")};


	$fh->print ("\t<mods:title>$title<\/mods:title>\n");
	if ($subtitle) 
		{$fh->print ("\t<mods:subTitle>$subtitle<\/mods:subTitle>\n");}
	$fh->print("<\/mods:titleInfo>\n\n");



### 2. MODS Name Element

	print "Before while loop: CurrentRow is $CurrentRow; LastRow is $LastRow; wfID is $wfID\n"; 

	my $namesToProcess="true";

	my $NextID = $Sheet->Range('A'.($CurrentRow+1))->{Value};

	while ($namesToProcess eq "true") 
		{

		if ($shortname && $given2)
			{
			$fh->print ("<mods:name type=\"personal\" authority=\"naf\">\n\t");
			$fh->print ("<mods:namePart type=\"family\">$family<\/mods:namePart>\n\t");
			$fh->print ("<mods:namePart type=\"given\">$given<\/mods:namePart>\n\t");
			$fh->print ("<mods:namePart type=\"given\">$given2<\/mods:namePart>\n\t");
			$fh->print ("<mods:displayForm>$family, $given $given2<\/mods:displayForm>\n\t");
			$fh->print ("<mods:affiliation>$dept, $school<\/mods:affiliation>\n\t");
			$fh->print ("<mods:role>\n\t\t<mods:roleTerm type=\"text\" authority=\"\">marcrelator<\/mods:roleTerm>\n\t<\/mods:role>\n\t");
			$fh->print ("<mods:description>$shortname<\/mods:description>\n");
			$fh->print ("<\/mods:name>\n");

			}

		if ($shortname && !$given2)
			{
			$fh->print ("<mods:name type=\"personal\" authority=\"naf\">\n\t");
			$fh->print ("<mods:namePart type=\"family\">$family<\/mods:namePart>\n\t");
			$fh->print ("<mods:namePart type=\"given\">$given<\/mods:namePart>\n\t");
			$fh->print ("<mods:displayForm>$family, $given<\/mods:displayForm>\n\t");
			$fh->print ("<mods:affiliation>$dept, $school<\/mods:affiliation>\n\t");
			$fh->print ("<mods:role>\n\t\t<mods:roleTerm type=\"text\" authority=\"marcrelator\">$marcRelatorCode<\/mods:roleTerm>\n\t<\/mods:role>\n\t");
			$fh->print ("<mods:description>$shortname<\/mods:description>\n");
			$fh->print ("<\/mods:name>\n");

			}
		if (!$shortname && $given2)
			{
			$fh->print ("<mods:name type=\"personal\">\n\t");
			$fh->print ("<mods:namePart type=\"family\">$family<\/mods:namePart>\n\t");
			$fh->print ("<mods:namePart type=\"given\">$given<\/mods:namePart>\n\t");
			$fh->print ("<mods:namePart type=\"given\">$given2<\/mods:namePart>\n\t");
			$fh->print ("<mods:displayForm>$family, $given $given2<\/mods:displayForm>\n\t");

			$fh->print ("<mods:role>\n\t\t<mods:roleTerm type=\"text\" authority=\"marcrelator\">$marcRelatorCode<\/mods:roleTerm>\n\t<\/mods:role>\n\t");
			$fh->print ("<\/mods:name>\n");

			}

		if (!$shortname && !$given2)
			{
			$fh->print ("<mods:name type=\"personal\">\n\t");
			$fh->print ("<mods:namePart type=\"family\">$family<\/mods:namePart>\n\t");
			$fh->print ("<mods:namePart type=\"given\">$given<\/mods:namePart>\n\t");
			$fh->print ("<mods:displayForm>$family, $given<\/mods:displayForm>\n\t");
			$fh->print ("<mods:role>\n\t\t<mods:roleTerm type=\"text\" authority=\"marcrelator\">$marcRelatorCode<\/mods:roleTerm>\n\t<\/mods:role>\n\t");
			$fh->print ("<\/mods:name>\n");

			}
		
		if ($Sheet->Range('A'.($CurrentRow+1))->{Value} && $wfID == $Sheet->Range('A'.($CurrentRow+1))->{Value})

			{
				print "next row is another one for this record\n";
				$row = shift @$usedRange;
				($wfID, $marcRelatorCode, $authorOrder, $family, $given, $given2, $shortname, $dept, $school, $title, $subtitle, $journalTitle, $enum1, $enum2, $chron2, $chron1, $startPage, $endPage, $pageList, $issn, $type, $url ,$doi, $setText, $ready, $version) = @$row;
				

				$CurrentRow++;
			
			}
		else
			{
				$namesToProcess="false";
			}
		}


### 3. MODS TypeOfResource Element

$fh->print("<mods:typeOfResource>text<\/mods:typeOfResource>\n\n");

### 4. MODS Genre Element


$fh->print("<mods:genre authority=\"marcgt\" type=\"workType\">$type<\/mods:genre>\n\n");
 
### 5. MODS OriginInfo Element

$fh->print("<mods:originInfo>\n");
	if ($chron1) {$fh->print("\t<mods:dateIssued>$chron1<\/mods:dateIssued>\n");}
	if ($chron1) {$fh->print("\t<mods:dateIssued encoding=\"w3cdtf\" keyDate=\"yes\">$chron1<\/mods:dateIssued>\n");}
	$fh->print("\t<mods:issuance>monographic<\/mods:issuance>\n");
$fh->print("<\/mods:originInfo>\n\n");

### 6.  MODS Language Element

$fh->print("<mods:language>\n\t<mods:languageTerm type=\"text\">English<\/mods:languageTerm>\n\t<mods:languageTerm type=\"code\" authority=\"iso639-2b\">eng<\/mods:languageTerm>\n<\/mods:language>\n\n");

### 7. MODS Physical Description

$fh->print("<mods:physicalDescription>\n");
	$fh->print("\t<mods:form authority=\"marcform\">electronic<\/mods:form>\n");
	$fh->print("\t<mods:internetMediaType>application/pdf<\/mods:internetMediaType>\n");
	$fh->print("\t<mods:digitalOrigin>reformatted digital<\/mods:digitalOrigin>\n");
$fh->print("<\/mods:physicalDescription>\n\n");

### 8. MODS Abstract


### 11. MODS Note Element
if ($setText) {$fh->print("\t<mods:note>$setText<\/mods:note>\n\n");}

if (($type ne "working paper") && $version==1)  
	{
		$fh->print("\t<mods:note type=\"version identification\">Version of record.<\/mods:note>\n\n")
	}

  if (($type ne "working paper") && $version==2)
	{
		$fh->print("\t<mods:note type=\"version identification\">Pre-print version of an article published in ")
	}

	if (($type ne "working paper") && $version==3)
	{
		$fh->print("\t<mods:note type=\"version identification\">Post-print version of an article published in ")
	};

	if (($type ne "working paper") && ($version==2||$version==3))
	{
		my $hostnonsort;
		if ($journalTitle =~ m/^The (.*)/) 
			{$hostnonsort = "The"; 
			$journalTitle=$1} 
		elsif ($journalTitle =~ m/^A (.*)/) 
			{$hostnonsort = "A";
			$journalTitle=$1} 
		elsif ($title =~ m /^An (.*)/) 
			{$hostnonsort = "An";
			$journalTitle=$1}; 
	
		if ($hostnonsort) {$fh->print ("$hostnonsort ")};

		$fh->print ("$journalTitle");

  
 		$fh->print(" $enum1");

		if ($enum2) {$fh->print ("\($enum2\)");};

		if ($endPage) {$fh->print (": $startPage-$endPage\.");}
		else {$fh->print (": $startPage\.");}
		if ($doi) {$fh->print(" doi:$doi\.");}
		$fh->print("<\/mods:note>\n\n");

		};





### 11. MODS Subject Element


### 14. MODS RelatedItem element

if ($version == 1)
	{
	$fh->print("<mods:relatedItem type=\"host\">\n\t<mods:titleInfo>");
	my $hostnonsort;
	if ($journalTitle =~ m/^The (.*)/) 
		{$hostnonsort = "The"; 
		$journalTitle=$1} 
	elsif ($journalTitle =~ m/^A (.*)/) 
		{$hostnonsort = "A";
		$journalTitle=$1} 
	elsif ($title =~ m /^An (.*)/) 
		{$hostnonsort = "An";
		$journalTitle=$1}; 
	
	if ($hostnonsort) {$fh->print ("\n\t\t<mods:nonSort>$hostnonsort <\/mods:nonSort>\n")};


	$fh->print ("\n\t\t<mods:title>$journalTitle<\/mods:title>\n");

$fh->print("\n\t<\/mods:titleInfo>\n");
	if ($issn) {$fh->print("\t<mods:identifier type=\"issn\">$issn<\/mods:identifier>\n");};
	  
 $fh->print("\t<mods:part>\n\t\t<mods:detail level=\"1\" type=\"volume\">\n\t\t              <mods:number>$enum1<\/mods:number>\n\t\t<\/mods:detail>\n");

if ($enum2) {$fh->print ("\t\t<mods:detail level=\"2\" type=\"issue\">\n\t\t               <mods:number>$enum2<\/mods:number>\n\t\t<\/mods:detail>\n");};

$fh->print ("\t\t<mods:extent unit=\"pages\">\n\t\t<mods:start>$startPage<\/mods:start>\n");

if ($endPage) {$fh->print ("\t\t\t<mods:end>$endPage<\/mods:end>\n\t\t\t<mods:list>pp. $startPage-$endPage<\/mods:list>\n");}
else {$fh->print ("\t\t\t<mods:list>p. $startPage<\/mods:list>\n");}

$fh->print ("\t\t</mods:extent>\n");

	if ($chron2){$fh->print("\t\t<mods:date>$chron2 $chron1<\/mods:date>\n\t<\/mods:part>\n");}
	else {$fh->print("\t\t<mods:date>$chron1<\/mods:date>\n\t<\/mods:part>\n");};


	$fh->print("<\/mods:relatedItem>\n");
	}

### 15. Mods Identifier
if ($doi && $version == 1) {$fh->print("\t<mods:identifier type=\"doi\">$doi<\/mods:identifier>\n\n");}
### 16. MODS Location Element

##if ($url) {
##	$fh->print("<mods:location>\n\t");
##	$fh->print("<mods:url displayLabel=\"Link to document\">$url<\/mods:url>\n\t");
##	$fh->print("<\/mods:location>\n");
##	};


### 16. MODS Access Condition
	$fh->print("<mods:accessCondition type=\"useAndReproduction\">These materials are made available for use in research, teaching and private study, pursuant to U.S. Copyright Law. The user must assume full responsibility for any use of the materials, including but not limited to, infringement of copyright and publication rights of reproduced materials. Any materials used for academic research or otherwise should be fully credited with the source. The publisher or original authors may retain copyright to the materials.<\/mods:accessCondition>\n");

### 19. MODS Extension Element

if ($url) {
	$fh->print("<mods:extension>\n\t");
	$fh->print("<ingestFile>$url<\/ingestFile>\n\t");

	$fh->print("<\/mods:extension>\n");
	};


### 20. MODS RecordInfo Element

$fh->print("<mods:recordInfo>\n");	
	$fh->print("\t<mods:recordContentSource>MChB<\/mods:recordContentSource>\n");


	$fh->print("\t<mods:languageOfCataloging>\n\t\t<mods:languageTerm type=\"text\">English<\/mods:languageTerm>\n\t\t<mods:languageTerm type=\"code\" authority=\"iso639-2b\">eng<\/mods:languageTerm>\n\t<\/mods:languageOfCataloging>\n\n");
$fh->print("<\/mods:recordInfo>\n");

### Close MODS Record

	$fh->print("<\/mods:mods>\n\n");

### Increment CurrentRow

$CurrentRow++;

};


#
$fh->print("<\/mods:modsCollection>\n");
$fh->close();

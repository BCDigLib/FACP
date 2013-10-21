#!C:/Perl/bin/perl -w
use strict;
use IO::File;
use File::Basename qw(basename);
use Win32::OLE qw(in with);
use Win32::OLE::Const 'Microsoft Excel';
use utf8;
use Cwd;
use XML::Simple;
use Data::Dumper;

main();

sub main 
{

	my ($wfID, $marcRelatorCode, $authorOrder, $family, $given, $given2, $shortname, $dept, $school, $title, $subtitle, $journalTitle, $enum1, $enum2, $chron2, $chron1, $startPage, $endPage, $pageList, $issn, $type, $url ,$doi, $setText, $ready, $version, $authors);

	my($worksheet_name, $Sheet, $excel_object) = setup_EXCEL_object(shift);

	my $fh=open_ouput_file($worksheet_name);

	my $project_type = project_type_determination();

	my $digital_origin = digital_origin_determination();

	my $data = read_faculty_names_xml();

	##read and process each row in the EXCEL file
	my $usedRange = $Sheet->UsedRange()->{Value};
			
		shift(@$usedRange);

		my $CurrentRow=2;


		while (my $row=shift @$usedRange)
		{


			$fh->print("<mods:mods>\n");
			if ($project_type eq "spreadsheet") 
			{
			($wfID, $authors, $title, $journalTitle, $issn, $enum1, $enum2, $chron1, $chron2, $startPage, $endPage, $doi, $setText) = @$row;
			mods_title($fh, $title, $subtitle);
			mods_name_element_spreadsheet($fh, $authors, $data);
			mods_type_of_resource($fh);
			mods_genre_spreadsheet($fh);
			mods_origin_info($fh, $chron1);
			mods_language($fh);
			mods_physical_description($fh, $digital_origin);
			mods_note($fh, $setText, '1', $doi, 'article', $journalTitle, $enum1, $enum2, $startPage, $endPage);
			mods_note_spreadsheet($fh);
			mods_related_item($fh, '1', $journalTitle, $issn, $enum1, $enum2, $chron1, $chron2, $startPage, $endPage);
			mods_access_condition($fh);
			mods_identifier($fh, $doi);
			mods_extension($fh, $doi);
			mods_record_info($fh);

			}
			else
			{
			($wfID, $marcRelatorCode, $authorOrder, $family, $given, $given2, $shortname, $dept, $school, $title, $subtitle, $journalTitle, $enum1, $enum2, $chron2, $chron1, $startPage, $endPage, $pageList, $issn, $type, $url ,$doi, $setText, $ready, $version) = @$row;
			mods_title($fh, $title, $subtitle);

			my $namesToProcess="true";
			

			while ($namesToProcess eq "true")
			{
				mods_name_element_database($fh, $wfID, $marcRelatorCode, $family, $given, $given2, $shortname, $dept, $school);
	
				if ($Sheet->Range('A'.($CurrentRow+1))->{Value} && $wfID == $Sheet->Range('A'.($CurrentRow+1))->{Value})
				{

					$row = shift @$usedRange;
					 ($wfID, $marcRelatorCode, $authorOrder, $family, $given, $given2, $shortname, $dept, $school, $title, $subtitle, $journalTitle, $enum1, $enum2, $chron2, $chron1, $startPage, $endPage, $pageList, $issn, $type, $url ,$doi, $setText, $ready, $version)= @$row;

		
					$CurrentRow++;	
				}

				else
				{			
					$namesToProcess="false";
					$CurrentRow++;
				}
			}


			mods_type_of_resource($fh);
			mods_genre_database($fh, $type);
			mods_origin_info($fh, $chron1);
			mods_language($fh);
			mods_physical_description($fh, $digital_origin);
			mods_note($fh, $setText, $version, $doi, $type, $journalTitle, $enum1, $enum2, $startPage, $endPage);
			mods_related_item($fh, $version, $journalTitle, $issn, $enum1, $enum2, $chron1, $chron2, $startPage, $endPage);
			mods_access_condition($fh);
			mods_extension($fh, $url);
			mods_record_info($fh);


			}

			$fh->print("<\/mods:mods>\n\n");
		};

	
	close_output_file ($fh);


};


### ### LIST OF MODS ELEMENTS


### MODS TitleInfo Element

sub mods_title
{
#Read a tab-delimited line of metadata and assign each element to an appropriately named variable
#
my $fh=shift;
my $title=shift;
my $subtitle=shift;

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

};



### See End of Document for MODS Author Element 



### MODS TypeOfResource Element

sub mods_type_of_resource
{
my $fh = shift;
$fh->print("<mods:typeOfResource>text<\/mods:typeOfResource>\n\n");

}



### MODS Genre Element

sub mods_genre_spreadsheet
{
my $fh = shift;

$fh->print("<mods:genre authority=\"marcgt\" type=\"workType\">article<\/mods:genre>\n\n");

}

sub mods_genre_database
{
my $fh = shift;
my $type = shift;

$fh->print("<mods:genre authority=\"marcgt\" type=\"workType\">$type<\/mods:genre>\n\n");

}



### MODS OriginInfo Element

sub mods_origin_info
{
my $fh = shift;
my $chron1 = shift;

$fh->print("<mods:originInfo>\n");
	if ($chron1) {$fh->print("\t<mods:dateIssued>$chron1<\/mods:dateIssued>\n");}
	if ($chron1) {$fh->print("\t<mods:dateIssued encoding=\"w3cdtf\" keyDate=\"yes\">$chron1<\/mods:dateIssued>\n");}
	$fh->print("\t<mods:issuance>monographic<\/mods:issuance>\n");
$fh->print("<\/mods:originInfo>\n\n");
}



### MODS Language Element

sub mods_language
{
my $fh = shift;

$fh->print("<mods:language>\n\t<mods:languageTerm type=\"text\">English<\/mods:languageTerm>\n\t<mods:languageTerm type=\"code\" authority=\"iso639-2b\">eng<\/mods:languageTerm>\n<\/mods:language>\n\n");

}



### MODS Physical Description

sub mods_physical_description
{
my $fh = shift;
my $digital_origin = shift;

$fh->print("<mods:physicalDescription>\n");
	$fh->print("\t<mods:form authority=\"marcform\">electronic<\/mods:form>\n");
	$fh->print("\t<mods:internetMediaType>application/pdf<\/mods:internetMediaType>\n");
	$fh->print("\t<mods:digitalOrigin>$digital_origin<\/mods:digitalOrigin>\n");
$fh->print("<\/mods:physicalDescription>\n\n");

};



### MODS Note Element

sub mods_note
{

my ($fh, $setText, $version, $doi, $type, $journalTitle, $enum1, $enum2, $startPage, $endPage)= @_;


if ($setText) {$fh->print("\t<mods:note>$setText<\/mods:note>\n\n");}

if (($type ne "working paper") && $version==1)  
	{
		$fh->print("\t<mods:note type=\"version identification\">Version of record.<\/mods:note>\n\n")			
	}

  if ($doi && $version==1)  
	{ 
		$fh->print("\t<mods:note>Also available on publisher's site: http://dx.doi.org/$doi<\/mods:note>\n\n");	
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
		elsif ($journalTitle =~ m /^An (.*)/) 
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


	if ($doi && ($version==2||$version==3))
	{
		$fh->print("\t<mods:note>Final published version is available at: http://dx.doi.org/$doi<\/mods:note>\n\n");
	}

}


sub mods_note_spreadsheet
{
my $fh = shift;

$fh->print("<mods:note>Bibliographic data derived from the Web of Science: http://thomsonreuters.com/web-of-science<\/mods:note>\n\n");

}


### MODS RelatedItem element

sub mods_related_item
{

my ($fh, $version, $journalTitle, $issn, $enum1, $enum2, $chron1, $chron2, $startPage, $endPage) = @_;

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
	elsif ($journalTitle =~ m /^An (.*)/) 
		{$hostnonsort = "An";
		$journalTitle=$1}; 
	
	if ($hostnonsort) {$fh->print ("\n\t\t<mods:nonSort>$hostnonsort <\/mods:nonSort>\n")};


	$fh->print ("\n\t\t<mods:title>$journalTitle<\/mods:title>\n");

$fh->print("\n\t<\/mods:titleInfo>\n");
	if ($issn) {$fh->print("\t<mods:identifier type=\"issn\">$issn<\/mods:identifier>\n");};
	  
 $fh->print("\t<mods:part>\n\t\t<mods:detail level=\"1\" type=\"volume\">\n\t\t<mods:number>$enum1<\/mods:number>\n\t\t<\/mods:detail>\n");

if ($enum2) {$fh->print ("\t\t<mods:detail level=\"2\" type=\"issue\">\n\t\t<mods:number>$enum2<\/mods:number>\n\t\t<\/mods:detail>\n");};

if ($startPage) {$fh->print ("\t\t<mods:extent unit=\"pages\">\n\t\t<mods:start>$startPage<\/mods:start>\n");

if ($startPage && $endPage) {$fh->print ("\t\t\t<mods:end>$endPage<\/mods:end>\n\t\t\t<mods:list>pp. $startPage-$endPage<\/mods:list>\n");}
else {$fh->print ("\t\t\t<mods:list>p. $startPage<\/mods:list>\n");}

$fh->print ("\t\t</mods:extent>\n");}

	if ($chron2){$fh->print("\t\t<mods:date>$chron2 $chron1<\/mods:date>\n\t<\/mods:part>\n");}
	else {$fh->print("\t\t<mods:date>$chron1<\/mods:date>\n\t<\/mods:part>\n");};


	$fh->print("<\/mods:relatedItem>\n");
	}

}



### MODS Access Condition

sub mods_access_condition
{

my $fh=shift;

	$fh->print("<mods:accessCondition type=\"useAndReproduction\">These materials are made available for use in research, teaching and private study, pursuant to U.S. Copyright Law. The user must assume full responsibility for any use of the materials, including but not limited to, infringement of copyright and publication rights of reproduced materials. Any materials used for academic research or otherwise should be fully credited with the source. The publisher or original authors may retain copyright to the materials.<\/mods:accessCondition>\n");

}



### MODS Identifier Element

sub mods_identifier
{
my ($fh, $doi) = @_;

	$fh->print("<mods:identifier type=\"doi\">$doi<\/mods:identifier>\n");
}



### MODS Extension Element

sub mods_extension
{
my ($fh, $file) = @_;

	$fh->print("<mods:extension>\n\t");
	$fh->print("<ingestFile>$file<\/ingestFile>\n\t");
	$fh->print("<\/mods:extension>\n");
}



### MODS RecordInfo Element

sub mods_record_info
{
my $fh = shift;

$fh->print("<mods:recordInfo>\n");	
	$fh->print("\t<mods:recordContentSource>MChB<\/mods:recordContentSource>\n");


	$fh->print("\t<mods:languageOfCataloging>\n\t\t<mods:languageTerm type=\"text\">English<\/mods:languageTerm>\n\t\t<mods:languageTerm type=\"code\" authority=\"iso639-2b\">eng<\/mods:languageTerm>\n\t<\/mods:languageOfCataloging>\n\n");
$fh->print("<\/mods:recordInfo>\n");


}



### MODS Name Element

sub mods_name_element_database
{

my ($fh, $wfID, $marcRelatorCode, $family, $given, $given2, $shortname, $dept, $school) = @_;

		if ($shortname && $given2)
			{
			$fh->print ("<mods:name type=\"personal\" authority=\"naf\">\n\t");
			$fh->print ("<mods:namePart type=\"family\">$family<\/mods:namePart>\n\t");
			$fh->print ("<mods:namePart type=\"given\">$given<\/mods:namePart>\n\t");
			$fh->print ("<mods:namePart type=\"given\">$given2<\/mods:namePart>\n\t");
			$fh->print ("<mods:displayForm>$family, $given $given2<\/mods:displayForm>\n\t");
			$fh->print ("<mods:affiliation>$dept, $school<\/mods:affiliation>\n\t");
			$fh->print ("<mods:role>\n\t\t<mods:roleTerm type=\"text\" authority=\"marcrelator\">$marcRelatorCode<\/mods:roleTerm>\n\t");
			if ($marcRelatorCode eq "Author") 
				{$fh->print ("\t<mods:roleTerm type=\"code\" authority=\"marcrelator\">aut<\/mods:roleTerm>");}
			$fh->print ("\n\t<\/mods:role>\n\t<mods:description>$shortname<\/mods:description>\n");
			$fh->print ("<\/mods:name>\n");

			}


		if ($shortname && !$given2)
			{
			$fh->print ("<mods:name type=\"personal\" authority=\"naf\">\n\t");
			$fh->print ("<mods:namePart type=\"family\">$family<\/mods:namePart>\n\t");
			$fh->print ("<mods:namePart type=\"given\">$given<\/mods:namePart>\n\t");
			$fh->print ("<mods:displayForm>$family, $given<\/mods:displayForm>\n\t");
			$fh->print ("<mods:affiliation>$dept, $school<\/mods:affiliation>\n\t");
			$fh->print ("<mods:role>\n\t\t<mods:roleTerm type=\"text\" authority=\"marcrelator\">$marcRelatorCode<\/mods:roleTerm>\n\t");
			if ($marcRelatorCode eq "Author") 
				{$fh->print ("\t<mods:roleTerm type=\"code\" authority=\"marcrelator\">aut<\/mods:roleTerm>");}
			$fh->print ("\n\t<\/mods:role>\n\t<mods:description>$shortname<\/mods:description>\n");
			$fh->print ("<\/mods:name>\n");
			}


		if (!$shortname && $given2)
			{
			$fh->print ("<mods:name type=\"personal\">\n\t");
			$fh->print ("<mods:namePart type=\"family\">$family<\/mods:namePart>\n\t");
			$fh->print ("<mods:namePart type=\"given\">$given<\/mods:namePart>\n\t");
			$fh->print ("<mods:namePart type=\"given\">$given2<\/mods:namePart>\n\t");
			$fh->print ("<mods:displayForm>$family, $given $given2<\/mods:displayForm>\n\t");

			$fh->print ("<mods:role>\n\t\t<mods:roleTerm type=\"text\" authority=\"marcrelator\">$marcRelatorCode<\/mods:roleTerm>\n\t");
			if ($marcRelatorCode eq "Author") 
				{$fh->print("\t<mods:roleTerm type=\"code\" authority=\"marcrelator\">aut<\/mods:roleTerm>");}
			$fh->print ("\n\t<\/mods:role>\n\t<\/mods:name>\n");

			}


		if (!$shortname && !$given2)
			{
			$fh->print ("<mods:name type=\"personal\">\n\t");
			$fh->print ("<mods:namePart type=\"family\">$family<\/mods:namePart>\n\t");
			$fh->print ("<mods:namePart type=\"given\">$given<\/mods:namePart>\n\t");
			$fh->print ("<mods:displayForm>$family, $given<\/mods:displayForm>\n\t");
			$fh->print ("<mods:role>\n\t\t<mods:roleTerm type=\"text\" authority=\"marcrelator\">$marcRelatorCode<\/mods:roleTerm>\n\t");
			if ($marcRelatorCode eq "Author") 
				{$fh->print("\t<mods:roleTerm type=\"code\" authority=\"marcrelator\">aut<\/mods:roleTerm>");}
			$fh->print ("\n\t<\/mods:role>\n\t<\/mods:name>\n");

			
			}

}

sub mods_name_element_spreadsheet
{
#Read a tab-delimited line of metadata and assign each element to an appropriately named variable
#
my $fh=shift;
my $authors = shift;
my $data = shift;
my $family;
my $given; 
my $given2;
my $dept;
my $school;

my @authors = split(/\s*;\s*/, $authors);


foreach (@authors) {


my $display_form = $_;
my ($family_name, $given_name) = split(/\s*,\s*/, $display_form);
if ($given_name) { $given_name =~ s/\s*$//;}
my $isBC='false';

###### attempt to use username

	foreach my $e (@{$data->{'facultyNames'}})  {

	if ($e->{'shortname'} && $e->{'shortname'} eq $display_form) {
		$isBC='true';

		if ($e->{'naf'} && $e->{'naf'}=~m/\d*/)
			{$fh->print ("<mods:name type=\"personal\" authority=\"naf\">\n\t");}
		else {$fh->print ("<mods:name type=\"personal\">\n\t");}

		$fh->print ("<mods:namePart type=\"family\">$e->{'family'}<\/mods:namePart>\n\t");
		$fh->print ("<mods:namePart type=\"given\">$e->{'given'}<\/mods:namePart>\n\t");
		$fh->print ("<mods:displayForm>$e->{'calc'}<\/mods:displayForm>\n\t");
		$fh->print ("<mods:affiliation>$e->{'DEPT'}, $e->{'SCHL_CD'}<\/mods:affiliation>\n\t");
		$fh->print ("<mods:role>\n\t\t<mods:roleTerm type=\"text\" authority=\"marcrelator\">Author<\/mods:roleTerm>\n\t\t");
		$fh->print ("<mods:roleTerm type=\"code\" authority=\"marcrelator\">aut<\/mods:roleTerm>\n\t<\/mods:role>\n\t");
		$fh->print ("<mods:description>$e->{'shortname'}<\/mods:description>\n");
		$fh->print ("<\/mods:name>\n\n");

		}	

}

if ( $display_form =~ m/\(BC\)/i )  {
	$isBC='true';

	$given_name =~ s/ \(BC\)//;
	$display_form =~ s/ \(BC\)//;
	
	$fh->print ("<mods:name type=\"personal\">\n\t");
	$fh->print ("<mods:namePart type=\"family\">$family_name<\/mods:namePart>\n\t");
	$fh->print ("<mods:namePart type=\"given\">$given_name<\/mods:namePart>\n\t");
	$fh->print ("<mods:displayForm>$display_form<\/mods:displayForm>\n\t");
	$fh->print ("<mods:affiliation>Boston College<\/mods:affiliation>\n\t");
	$fh->print ("<mods:role>\n\t\t<mods:roleTerm type=\"text\" authority=\"marcrelator\">Author<\/mods:roleTerm>\n\t\t<mods:roleTerm type=\"code\" authority=\"marcrelator\">aut<\/mods:roleTerm>\n\t<\/mods:role>\n\t");
	$fh->print ("<mods:description>nonfaculty<\/mods:description>\n");
	$fh->print ("<\/mods:name>\n\n");

	}

if ($isBC eq 'false')  {

$fh->print ("<mods:name type=\"personal\">\n\t<mods:namePart type=\"family\">$family_name<\/mods:namePart>\n\t<mods:namePart type=\"given\">$given_name<\/mods:namePart>\n\t<mods:displayForm>$display_form<\/mods:displayForm>\n\t<mods:role>\n\t\t<mods:roleTerm type=\"text\" authority=\"marcrelator\">Author<\/mods:roleTerm>\n\t\t<mods:roleTerm type=\"code\" authority=\"marcrelator\">aut<\/mods:roleTerm>\n\t<\/mods:role>\n<\/mods:name>\n\n");};

	} 


};


### ### OTHER TASKS


###  Open and Setup Excel


sub setup_EXCEL_object {

#Get the name of the excel workbook and worksheet you want to process
print "\n\nEnter the name of the Excel file containing \nthe data you wish to convert to MODS: ";
my $excelfile = <STDIN>; 
chomp $excelfile; 
exit 0 if (!$excelfile);

print "\n\nName of the worksheet containing the \ndata you wish to convert to MODS: ";
my $worksheet_name = <STDIN>; 
chomp $worksheet_name; 
exit 0 if (!$worksheet_name);

my $dir = getcwd;
$dir=~s/\//\\/g;
#print "dir is $dir\n";
$excelfile=$dir."\\".$excelfile;

#Get Ready to use $Win32::OLE

$Win32::OLE::Warn = 3; # Die on Errors.

# ::Warn = 2; throws the errors, but #
# expects that the programmer deals  #

#Create an EXCEL object to work with and define how the object is going to exit

my $Excel = Win32::OLE->GetActiveObject('Excel.Application')
        || Win32::OLE->new('Excel.Application', 'Quit');

#Turn off all the alert boxes, such as the SaveAs response "This file already exists", etc. using the DisplayAlerts property.

$Excel->{DisplayAlerts}=0;   

#Open an existing file to work with 
                                                 
my $book_object = $Excel->Workbooks->Open($excelfile);   

#Create a reference to a worksheet object and activate the sheet to give it focus so that actions taken on the workbook or application objects occur on this sheet unless otherwise specified.

my $sheet_object = $book_object->Worksheets($worksheet_name);
$sheet_object->Activate();  

return ($worksheet_name, $sheet_object, $Excel);
}



### Open Output File and Print XML declaration and root node

sub open_ouput_file {

my $fh=shift;

$fh = IO::File->new($fh.'.xml', 'w')
	or die "unable to open output file for writing: $!";
binmode($fh, ':utf8');
$fh->print("<?xml version='1.0' encoding='UTF-8' ?>\n");
$fh->print("<mods:modsCollection xmlns:xlink=\"http://www.w3.org/1999/xlink\" xmlns:xsi=\"http://www.w3.org/2001/XMLSchema-instance\" xmlns:mods=\"http://www.loc.gov/mods/v3\" xsi:schemaLocation=\"http://www.loc.gov/mods/v3 http://www.loc.gov/standards/mods/v3/mods-3-4.xsd\">\n");

return($fh);

};



### Determine Whether This is a Database or Spreadsheet Project

sub project_type_determination
{
	print "\n\nProject type: eScholarship database export \nor pdf project spreadsheet? \nEnter database or spreadsheet: ";
	my $project_type = <STDIN>; 
	chomp $project_type;
	exit 0 if ($project_type ne "database" && $project_type ne "spreadsheet"); 
	return ($project_type);
}



### Determine Digital Origin 

sub digital_origin_determination
{
	print "\n\nWhat is the digital origin of this stuff?: ";
	my $digital_origin = <STDIN>; 
	chomp $digital_origin;
	exit 0 if ($digital_origin ne "born digital" && $digital_origin ne "reformatted digital"); 
	return ($digital_origin);
}



### Read facultyNames.xml

sub read_faculty_names_xml
{

# create object
my $xml = new XML::Simple;

# read XML file
my $data = $xml->XMLin("facultyNames.xml");

#commenting this block out, cause we've already proved PERL is reading the xml file from ACCESS
#use Data Dumper to confirm xml file was read into perl
#print Dumper($data);  

return($data);

};



### Close Output File

sub close_output_file{
my $fh=shift;
$fh->print("<\/mods:modsCollection>\n");
$fh->close();

};

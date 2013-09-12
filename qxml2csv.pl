#!/usr/bin/perl
#
#	qxml2csv.pl
#
#	From CLI take a Qualys XML output file and shit out CSV of the hosts and vulns in the file
#	Redirect > somfile.csv
#
#	Needs to contain risks, results, description, and remediations
#
#	becomes qxmlproc.pl
#
use XML::Simple;
require 'dumpvar.pl';
use Spreadsheet::Writeexcel;



my $HeaderRow  = "IP Address,DNS Name,NETBIOS Name,Asset Groups,";
$HeaderRow    .= "QID,First Found,Last Found,Times Found,Vulnerability Status,Port,Protocol,";
$HeaderRow    .= "Vulnerability Title,Severity,CVSS Base,CVSS Temporal,Threat,Impact,Solution";

my @header = split /\,/, $HeaderRow;


#        <THREAT><![CDATA[The rex daemon is an RPC program that enables unauthorized remote users to execute commands without a password.]]></THREAT>
#        <IMPACT><![CDATA[Unauthorized users can execute commands as root from a remote system (no authentication is required).]]></IMPACT>
#        <SOLUTION><![CDATA[Running this RPC daemon on your server creates a severe vulnerability.  Remove it from the list of RPC programs to be loaded at start up. On SunOS, this program is usually located in the &quot;/etc/init.d/rpc&quot; file.]]></SOLUTION>



my $infile = shift;
#my $infile = "Report.xml";
my @TheReallyBigList = ();
my %QIDInfoList = ();
my %AGList = ();


my $xs = XML::Simple->new();
my $reff = $xs->XMLin($infile);
#dumpValue ($reff);

foreach $thing (keys %{$reff}) {
	#print "A thing: '$thing'  It is a '$reff->{$thing}'\n";

	if ($thing eq "HEADER") {
		#dumpValue ($reff->{$thing});
	}
	
	if ($thing eq "GLOSSARY") {
		#dumpValue ($reff->{$thing});
		$qidinfo = $reff->{$thing}{VULN_DETAILS_LIST}{VULN_DETAILS};
		foreach $qidkey ( keys %{$qidinfo} ) {
			#print "A QID: $qidkey\n";
			#dumpValue ($qidinfo->{$qidkey});print "\n";
			
			my $title = $qidinfo->{$qidkey}{TITLE};
			$title =~ s/\,//g;
			my $threat = $qidinfo->{$qidkey}{THREAT};
			$threat =~ s/\"//g;
			$threat =~ s/\,//g;
			$threat =~ s/\t/\./g;
			$threat =~ s/\<PRE\>//g;
			$threat =~ s/\<\/PRE\>//g;
			$threat =~ s/\&quot\;/\"/g;
			$threat =~ s/\<A HREF\=/ /g;
			$threat =~ s/TARGET\=.*\<\/A\>/\n/g;
			$threat =~ s/\n//g;
			$threat =~ s/\&gt\;/\>/g;
			$threat =~ s/\<LI\>/\n/g;
			$threat =~ s/\<BR\>/\n/g;
			$threat =~ s/\<P\>/\n/g;
			$threat =~ s/         //g;
			my $impact = $qidinfo->{$qidkey}{IMPACT};
			$impact =~ s/\"//g;
			$impact =~ s/\,//g;
			$impact =~ s/\t/\./g;
			$impact =~ s/\<PRE\>//g;
			$impact =~ s/\<\/PRE\>//g;
			$impact =~ s/\&quot\;/\"/g;
			$impact =~ s/\<A HREF\=/ /g;
			$impact =~ s/TARGET\=.*\<\/A\>/\n/g;
			$impact =~ s/\n//g;
			$impact =~ s/\&gt\;/\>/g;
			$impact =~ s/\<LI\>/\n/g;
			$impact =~ s/\<BR\>/\n/g;
			$impact =~ s/\<P\>/\n/g;
			$impact =~ s/         //g;
			my $solution = $qidinfo->{$qidkey}{SOLUTION};
			$solution =~ s/\"//g;
			$solution =~ s/\,//g;
			$solution =~ s/\t/\./g;
			$solution =~ s/\<PRE\>//g;
			$solution =~ s/\<\/PRE\>//g;
			$solution =~ s/\&quot\;/\"/g;
			$solution =~ s/\<A HREF\=/ /g;
			$solution =~ s/TARGET\=.*\<\/A\>/\n/g;
			$solution =~ s/\n//g;
			$solution =~ s/\&gt\;/\>/g;
			$solution =~ s/\<LI\>/\n/g;
			$solution =~ s/\<BR\>/\n/g;
			$solution =~ s/\<P\>/\n/g;
			$solution =~ s/         //g;
			my $severity = $qidinfo->{$qidkey}{SEVERITY};
			my $lastupdate = $qidinfo->{$qidkey}{LAST_UPDATE};
			my $category = $qidinfo->{$qidkey}{CATEGORY};
			
			my $cvss_base,$cvss_temp;
			if ( ref($qidinfo->{$qidkey}{CVSS_SCORE}{CVSS_BASE}) eq "HASH" ) {
				$cvss_base = $qidinfo->{$qidkey}{CVSS_SCORE}{CVSS_BASE}{content};
			} else {
				$cvss_base = $qidinfo->{$qidkey}{CVSS_SCORE}{CVSS_BASE};
			}
			$cvss_temp = $qidinfo->{$qidkey}{CVSS_SCORE}{CVSS_TEMPORAL};
			
			$QIDInfoList{$qidkey} = "\"$title\",$severity,$cvss_base,$cvss_temp,$threat,$impact,$solution";
			#$QIDInfoList{$qidkey} = "$title,$severity,$cvss_base,$cvss_temp";
			
			#print "INFO: $QIDInfoList{$qidkey}\n";
			
		}
	}

	if ($thing eq "APPENDICES") {
		#dumpValue ($reff->{$thing});
	}	
	
	if ($thing eq "HOST_LIST") {
		foreach $thing2 ( keys %{$reff->{$thing}} ) {
			#print "A thing2: '$thing2'  It is a '$reff->{$thing}{$thing2}'\n";
			
			foreach $listitem ( @{$reff->{$thing}{$thing2}} ) {
				#print "A listitem: '$listitem'\n";
				
				my $ip = $listitem->{IP};
				my $os = $listitem->{OPERATING_SYSTEM};
				my $trackingmethod = $listitem->{TRACKING_METHOD};
				my $dns = $listitem->{DNS};
				my $netbios = $listitem->{NETBIOS};
				
				# Asset Groups...
				my @agfixed = ();
				my $ags;
				my $ag = $listitem->{ASSET_GROUPS}{ASSET_GROUP_TITLE};
				if (ref($ag) eq "ARRAY") {
					foreach $group (@{$ag}) {
						#print "An asset group: $group\n";
						if ($group eq "mgt_PCI_All") { next; }
						if ($group =~ /^mgt_PCI_(.*)/i) {
							push @agfixed, $1;
							#print "ADDING AG: $1\n";
							$AGList{$1}++;
						}
					}
					$ags = join ':', @agfixed;
				} else {
					$ags = $ag;
					$ags =~ s/^mgt_PCI_(.*)/$1/i;
					$AGList{$ags}++;
				}
				
				
				
				#print "$ip,$dns,$netbios,$ags,VULN!\n";
				
				my $hostinfo = "$ip,$dns,$netbios,$ags";
				
				if ( ref($listitem->{VULN_INFO_LIST}{VULN_INFO}) eq "ARRAY" ) {

					foreach $vuln ( @{$listitem->{VULN_INFO_LIST}{VULN_INFO}} ) {
						
						#dumpValue ($vuln);
						#print "\n";
						
						my $qid = $vuln->{QID}{id};
						
						# 2008-10-12T07:06:02Z  ==>  2008-10-12
						my $firstfound = $vuln->{FIRST_FOUND};
						my $lastfound = $vuln->{LAST_FOUND};
						$firstfound =~ s/^(\d\d\d\d-\d\d-\d\d)T\d\d:\d\d:\d\dZ/$1/;
						$lastfound =~ s/^(\d\d\d\d-\d\d-\d\d)T\d\d:\d\d:\d\dZ/$1/;
						
						
						
						my $ssl = $vuln->{SSL};
						my $timesfound = $vuln->{TIMES_FOUND};
						my $type = $vuln->{TYPE};
						my $vulnstatus = $vuln->{VULN_STATUS};
						my $port = $vuln->{PORT};
						my $protocol = $vuln->{PROTOCOL};
						#my $result = $vuln->{RESULT};
						
						
						my $vulninfo = "$qid,$firstfound,$lastfound,$timesfound,$vulnstatus,$port,$protocol";
						#print "$hostinfo,$vulninfo\n";
						#print "SOME QID Goodness for $qid: '$QIDInfoList{$qid}'\n";
						push @TheReallyBigList, "$hostinfo,$vulninfo";
						
					}
				} elsif (ref($listitem->{VULN_INFO_LIST}{VULN_INFO}) eq "HASH") {
					#print "A VULN HASH??? What is this? Here:\n";
					#dumpValue ($listitem);
					my $vuln = $listitem->{VULN_INFO_LIST}->{VULN_INFO};
					#dumpValue ($vuln); print "\n";
					#my $v2 = $vuln->{VULN_INFO};
					#dumpValue ($v2); print "\n";
					
					my $qid = $vuln->{QID}{id};
					
					# 2008-10-12T07:06:02Z  ==>  2008-10-12
					my $firstfound = $vuln->{FIRST_FOUND};
					my $lastfound = $vuln->{LAST_FOUND};
					$firstfound =~ s/^(\d\d\d\d-\d\d-\d\d)T\d\d:\d\d:\d\dZ/$1/;
					$lastfound =~ s/^(\d\d\d\d-\d\d-\d\d)T\d\d:\d\d:\d\dZ/$1/;
					
					
					
					
					my $ssl = $vuln->{SSL};
					my $timesfound = $vuln->{TIMES_FOUND};
					my $type = $vuln->{TYPE};
					my $vulnstatus = $vuln->{VULN_STATUS};
					my $port = $vuln->{PORT};
					my $protocol = $vuln->{PROTOCOL};
					#my $result = $vuln->{RESULT};

					
					my $vulninfo = "$qid,$firstfound,$lastfound,$timesfound,$vulnstatus,$port,$protocol";
					#print "$hostinfo,$vulninfo\n";
					#print "SOME QID Goodness for $qid: '$QIDInfoList{$qid}'\n";
					push @TheReallyBigList, "$hostinfo,$vulninfo";
					
				} else {
					print "WELL NOW WHAT THE FUCK ARE YOU?\n";
					dumpValue ($listitem);
				}
				
				
				
				#print "\n";
				
			}
		}
	}
}

#dumpValue (\%QIDInfoList);
#foreach $t (keys %QIDInfoList) { print "$t\n"; }
#exit();


# Print TheReallyBigList to <infile>.csv
my $outfilecsv = $infile;
($outfilecsv) =~ s/(.*)\.xml/$1\.csv/;
open OUT, ">$outfilecsv";
print "writing $outfilecsv\n";

print OUT "$HeaderRow\n";
foreach $line (@TheReallyBigList) {
	my ($thisqid) = $line =~ /.*(qid_\d+).*/;
	#print "\n\n$thisqid\n\n";
	print OUT "$line,$QIDInfoList{$thisqid}\n";
}
close OUT;

# Now dump TheReallyBigList into the first (SUMMARY) tab of an excel workbook...
my $outfilexls = $infile;
($outfilexls) =~ s/(.*)\.xml/$1\.xls/;
open OUT, "$outfilexls";
print "writing $outfilexls\n";

my $wb = Spreadsheet::WriteExcel->new($outfilexls);
my $ws = $wb->add_worksheet('SUMMARY');

my $header_fmt = $wb->add_format();
$header_fmt->set_bold();
$header_fmt->set_align('left');
$header_fmt->set_bg_color(50); # light green

my $row = 0;
$ws->write($row, 0, \@header, $header_fmt);
foreach $line (@TheReallyBigList) {
	my ($thisqid) = $line =~ /.*(qid_\d+).*/;
	#print "\n\n$thisqid\n\n";
	@data = split /\,/, "$line,$QIDInfoList{$thisqid}";
	
	$ws->write($row++, 0, \@data);
}








#foreach $assetgroup (keys %AGList) {
#	$row = 0;
#	
#	my $agtab = $assetgroup;
#	$agtab =~ s/\//-/;
#	my $ws2 = $wb->add_worksheet($agtab);
#	
#	print "$assetgroup => $agtab\n";
#
#	$ws2->write($row, 0, \@header, $header_fmt);
#	
#	foreach $line (@TheReallyBigList) {
#		my ($thisqid) = $line =~ /.*(qid_\d+).*/;
#		#print "\n\n$thisqid\n\n";
#		@data = split /\,/, "$line,$QIDInfoList{$thisqid}";
#		
#		if ($line =~ /$assetgroup/) {
#			$ws2->write($row++, 0, \@data);
#		}
#	}
#	
#
#
#
#}















#$ws->write($row, \@header, $header_fmt);



#    # Create a new Excel workbook
#    my $workbook = Spreadsheet::WriteExcel->new('perl.xls');
#
#    # Add a worksheet
#    $worksheet = $workbook->add_worksheet();
#
#    #  Add and define a format
#    $format = $workbook->add_format(); # Add a format
#    $format->set_bold();
#    $format->set_color('red');
#    $format->set_align('center');
#
#    # Write a formatted and unformatted string, row and column notation.
#    $col = $row = 0;
#    $worksheet->write($row, $col, 'Hi Excel!', $format);
#    $worksheet->write(1,    $col, 'Hi Excel!');
#
#    # Write a number and a formula using A1 notation
#    $worksheet->write('A3', 1.2345);
#    $worksheet->write('A4', '=SIN(PI()/4)');



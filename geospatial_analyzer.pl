use Parse::CSV;
use Config::Tiny;
use Data::Dumper;
use DateTime;  
use Time::Piece;
use Spreadsheet::Read qw(ReadData rows row cellrow);
use LWP::Simple;
use JSON::Parse qw(parse_json assert_valid_json);
use Excel::Writer::XLSX;
use constant CONFIDENCE => [ "Low","Medium","High" ];

# start the stopwatch
my $total_time_start = time();
# set up bing config
my $config = Config::Tiny->read( "config.ini", 'utf8' );
my $geospatial_url = $config->{API}{geospatial_url};
my $geocode_url = $config->{API}{geocode_url};
my $api_key = $config->{API}{key};
my $curate_polys = $config->{API}{curate_polys};
my $response_format = $config->{API}{format};
my $county_entity_type = $config->{API}{county_entity_type};
my $muni_entity_type = $config->{API}{muni_entity_type};
my $get_all = $config->{API}{get_all};
my $get_metadata = $config->{API}{get_metadata};
my $culture = $config->{API}{culture};
my $user_region = $config->{API}{user_region};
my $spatial_filter = $config->{API}{spatial_filter};
my %confidence_rating = (Low => 1, Medium => 2, High => 3);
# coordinates below this confidence rating are removed from results
my $minimum_confidence = $confidence_rating{$config->{API}{minimum_confidence}};
# if this setting is true, matching localities between the input address and output address 
# during geocoding boost the confidence rating by one point
my $match_confidence_adjustment = $config->{API}{locality_match_confidence_adjustment};
# if this setting has a value, it will guide the geocoder
my $user_location = $config->{API}{user_location};
# set up args
my $arg_count = $#ARGV + 1;
unless($arg_count == 6){
    die('usage: geospatial_analyzer data_file out_file address_col address_2_col locality_col state_col');
}
# read in input
my $data_file = Spreadsheet::Read->new ($ARGV[0]);
my $info     = $data_file->[0];
my $sheet = $data_file->sheet (1);
print "Parsed $ARGV[0] with $info->{parser}-$info->{version}\n";
print "Sheet 1 has ".$sheet->maxcol." columns and ".$sheet->maxrow." rows\n";

my $address_col = $ARGV[2];
my $address_2_col = $ARGV[3];
my $locality_col = $ARGV[4];
my $state_col = $ARGV[5];

# create output
my $workbook = Excel::Writer::XLSX->new( $ARGV[1] );
my $worksheet = $workbook->add_worksheet();

my %geoinfo;

my @headers = $sheet->row(1);
push @headers, 'Within County';
push @headers, 'Within Municipality';
my $h_ref = \@headers;
 
$worksheet->write_row( 0, 0, $h_ref );
my $last_percent = 0;

# start at 2 to skip header row $sheet->maxrow
print "Starting to process...\n";
foreach my $i (2 .. $sheet->maxrow) {
    my @row_data = $sheet->row($i);
    my $state = $sheet->cell ($state_col, $i);
    my $locality = $sheet->cell ($locality_col, $i);
    my $address = $sheet->cell ($address_col, $i);
    #my $address = $sheet->cell ($address_col, $i)." ".$sheet->cell ($address_2_col, $i);
    # we use full address as the key into the hash because there could be multiple towns with the same street names.
    # replace periods cause they mess up the API      
    $address =~ s/\x2E/ /ig;
    # replace hashes cause they mess up the API
    $address =~ s/\x23/ /ig;
    # replace question marks cause they really mess up the API
    $address =~ s/\x3F/ /ig;
    $address =~ s/^\s+|\s+$//g;
    # replace periods cause they mess up the API      
    $locality =~ s/\x2E/ /ig;
    $locality =~ s/^\s+|\s+$//g;
    
    #use address with locality in case two localities have the same street
    my $key = uc("$locality$address");
    $key =~ s/\s+//g;

    # skip API calls for duplicate addresses 
    unless (exists($geoinfo{$key}))
    {   
        my $coord_url = $geocode_url."/$state/$locality/$address?";
        unless($user_location == undef || $user_location eq ''){
            $coord_url = $coord_url."&userLocation=$user_location";
            $coord_url = $coord_url."&key=$api_key";  
        }
        else{
            $coord_url = $coord_url."key=$api_key";  
        }
              
        my $coords = undef;
        # first get request retrieves longitude and latitude of address
         print("$coord_url\n");
        eval {
            my $content = get($coord_url);
            my $coord_data = parse_json ($content);
            $confidence = $confidence_rating{$coord_data->{'resourceSets'}[0]->{'resources'}[0]->{'confidence'}};
            # adjust confidence based on parts of the output that match the input
            if($match_confidence_adjustment){
                $output_locality = lc($coord_data->{'resourceSets'}[0]->{'resources'}[0]->{'address'}{'locality'});
                if($output_locality eq lc($locality)){
                    $confidence += 1;
                }
            }
            if($confidence >= $minimum_confidence){
                $coords = $coord_data->{'resourceSets'}[0]->{'resources'}[0]->{'point'}{'coordinates'};
            }
            else{
                die("confidence too low to geocode\n");
            }
        };
        if($@){
            print("$@\n");
            print("could not geocode address\n"); 
        }

        unless($coords == undef){
            my $long = $coords->[0];
            my $lat = $coords->[1];
            my $county_url = $geospatial_url."?SpatialFilter=$spatial_filter($long,$lat,0,'$county_entity_type',$get_all,$get_metadata,'$culture','$user_region')&PreferCuratedPolygons=$curate_polys&\$format=$response_format&key=$api_key";
            my $muni_url = $geospatial_url."?SpatialFilter=$spatial_filter($long,$lat,0,'$muni_entity_type',$get_all,$get_metadata,'$culture','$user_region')&PreferCuratedPolygons=$curate_polys&\$format=$response_format&key=$api_key"; 
            $geoinfo{$key}{'municipality'} = undef;
            $geoinfo{$key}{'county'} = undef;
            my $content = "";
            eval {
                $content = get($muni_url);
                my $muni_data = parse_json ($content);
                $geoinfo{$key}{'municipality'} = $muni_data->{d}->{results}[0]->{'Name'}{'EntityName'};
            };
            if($@){      
                print("$@\n");       
                print("No municipality was found\n");       
            }
            
            eval {
                $content = get($county_url);    
                my $county_data = parse_json ($content);
                $geoinfo{$key}{'county'} = $county_data->{d}->{results}[0]->{'Name'}{'EntityName'};
            };
            if($@){
                print("$@\n"); 
                print("No county was found\n");
            }
        }
    }
    push @row_data, $geoinfo{$key}{'county'};
    push @row_data, $geoinfo{$key}{'municipality'};
    my $row_data_ref = \@row_data;
    $worksheet->write_row($i-1, 0, $row_data_ref );
    my $percent = ($i/$sheet->maxrow)*100;
    if(($last_percent == 0 && $percent >= 1.0) || $percent >= $last_percent+5){
        my $current_time = (time() - $total_time_start);
        my $time_piece = "seconds";
        if($current_time > 60.0){
        $current_time /= 60.0;
        $time_piece = "minutes"
        }
        printf("%.0f%% complete at %.2f $time_piece\n",$percent,$current_time);
        $last_percent = $percent;
    }
    
}
$workbook->close(); 
my $unique_count = scalar keys %geoinfo;
my $total_transactions = 3 * $unique_count;
print("Gecoded $unique_count unique addresses with $total_transactions total transactions.\n");
#log completion time   
my $complete_time = (time() - $total_time_start)/60.0;
print("script finished in $complete_time minutes\n");

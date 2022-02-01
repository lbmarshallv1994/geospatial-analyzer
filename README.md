This PERL program uses the **Bing Maps API** to figure out which city or county an address is within the limits of. The program takes in an excel spreadsheet containing the address, city and county and outputs a similar excel sheet with "Within County" and "Within Municipality" appended.

## Dependencies

This program uses the following PERL dependencies:
* Parse::CSV
* Config::Tiny
* Data::Dumper
* DateTime
* Time::Piece
* Spreadsheet::Read
* LWP::Simple
* JSON::Parse
* Excel::Writer::XLSX

These should all be on CPAN.

## Installation

1. Create your account with Microsoft and apply for a key with the Bing Maps API, it's free to use with limitations.
2. Set up your API connection in **config.ini** using the provided example file.
    * enter the Key Microsoft gives you at **key=**
    * enter your longitude and lattitude into **user_location=** this will increase the accuracy of the analyzer's findings 
3. Organize your data into an excel sheet, Address line 1, line 2, city, county, state and post code should be **separate columns**.
4. Install the dependencies using CPAN.
    ```
    cpan -i Parse::CSV
    ```
## Usage

```
perl geospatial_analyzer.pl data_file out_file address_col address_2_col locality_col state_col
```

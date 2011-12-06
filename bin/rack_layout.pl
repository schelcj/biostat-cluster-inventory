#!/usr/bin/perl

use FindBin qw($Bin);
use Modern::Perl;
use Spreadsheet::ParseExcel;
use Excel::Template;
use Data::Dumper;

my %racks      = ();
my $sheet_name = q{Equipment};
my $inventory  = qq{$Bin/../data/inventory.xls};
my $col_map    = {host => 0, rack => 6, rack_u => 7, rack_loc => 8};

my $parser   = Spreadsheet::ParseExcel->new();
my $workbook = $parser->parse($inventory);
my $sheet    = $workbook->worksheet($sheet_name);

my ($min, $max) = $sheet->row_range();
for my $row (($min + 1) .. $max) {
  next if not $sheet->get_cell($row, '0');
  my %entry = ();

  for my $col (keys %{$col_map}) {
    $entry{$col} = $sheet->get_cell($row, $col_map->{$col})->value();
  }

  my $rack = delete $entry{rack};

  if ($entry{rack_u} =~ /(\d+)\-(\d+)/) {
    my $start = $1;
    my $end   = $2;

    for my $u ($start .. $end) {
      my %dupe = %entry;
      $dupe{rack_u} = $u;
      push @{$racks{$rack}}, \%dupe;
    }
  } else {
    push @{$racks{$rack}}, \%entry;
  }
}

for my $rack (keys %racks) {
  my $excel  = Excel::Template->new(filename => qq{$Bin/../templates/rack_layout.xml});
  my $report = qq{$Bin/../reports/rack_layout-$rack.xls};
  my $params = {results => [], rack => $rack,};

  for my $item (sort {$b->{rack_u} <=> $a->{rack_u}} @{$racks{$rack}}) {
    push @{$params->{'results'}}, $item;
  }

  $excel->param($params);
  $excel->write_file($report);
}

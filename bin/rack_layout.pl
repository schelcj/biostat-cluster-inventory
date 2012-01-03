#!/usr/bin/perl

## no critic (ValuesAndExpressions::ProhibitMagicNumbers)
#
use FindBin qw($Bin);
use Modern::Perl;
use Spreadsheet::ParseExcel;
use Excel::Template;

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
    my $cell = $sheet->get_cell($row, $col_map->{$col});

    if (defined $cell) {
      $entry{$col} = $cell->value();
    }
  }

  my $rack = delete $entry{rack};
  next if not $rack;

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

for my $name (keys %racks) {
  my $excel  = Excel::Template->new(filename => qq{$Bin/../templates/rack_layout.xml});
  my $report = qq{$Bin/../reports/rack_layout-$name.xls};
  my $params = {results => [], rack => $name,};
  my @items  = map {rack_u => $_, rack_loc => 'empty', host => 'empty', empty => 1}, (1 .. 42);

  for my $item (@{$racks{$name}}) {
    $items[$item->{rack_u} - 1] = $item;
  }

  @items = reverse sort {$a->{rack_u} <=> $b->{rack_u}} @items;

  $params->{results} = \@items;

  $excel->param($params);
  $excel->write_file($report);
}

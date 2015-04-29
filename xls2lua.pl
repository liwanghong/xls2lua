#!/usr/bin/perl

use Spreadsheet::Read;
use Spreadsheet::XLSX;
use Getopt::Long;
use Data::Dumper; 
 
my $xls_file;
my $opt;
my $id_col = 0;
my $num2_str = 0;
GetOptions("x|xls=s" => \$xls_file,
           "o|opt=s" => \$opt,
           "c|col:i" => \$id_coli,
           "num|n"=>\$num2_str)
       or die ("Error param");

if ($opt ne "a" and $opt ne "h")
{
    die ("Invalid opt\n");
}

my $book = Spreadsheet::XLSX->new("$xls_file");
my @titles;
foreach my $sheet (@{$book->{Worksheet}})
{
    my $row = $sheet->{MaxRow};
    my $col = $sheet->{MaxCol};
    my @titles = ();
    for (my $col_index = 0; $col_index <= $col; $col_index = $col_index + 1)
    {
        my $cell = $sheet->{Cells}[0][$col_index];
        my $value = $cell->{Val};
        push(@titles, $value);
    }
    print "$sheet->{Name} = $sheet->{Name} or {\n";
    for(my $row_index = 1; $row_index <= $row; $row_index = $row_index + 1)
    {
        my $key_cell = $sheet->{Cells}[$row_index][$id_col];
        my $key = $key_cell->{Val};
        if ($opt eq "a")
        {
            print "    {\[\"$titles[$id_col]\"\] = \"$key\", ";
        }
        elsif ($opt eq "h")
        {
            print "    [\"$key\"] = {";
        }
        for (my $col_index = $id_col + 1; $col_index <= $col; $col_index = $col_index + 1)
        {
    	    my $cell = $sheet->{Cells}[$row_index][$col_index];
            my $value = $cell->{Val};
            $value =~ s/^\s+//;
            $value =~ s/\s+$//;
            if ($num2_str)
            {
                if($value =~ /^\s*$/)
                {
                    print "\[\"$titles[$col_index]\"\] = nil";
                }
                elsif ($value =~ /^[\+-\d\.,{}\s]+$/) 
                {
                    print "\[\"$titles[$col_index]\"\] = $value";
                }
                else
                {
                    print "\[\"$titles[$col_index]\"\] = \"$value\"";
                }
            }
            else
            {
                print "\[\"$titles[$col_index]\"\] = \"$value\"";
            }
            if ($col_index != $col)
            {
                print ", ";
            }
        }
        if ($row_index == $row)
        {
            print "}\n";
        }
        else
        {
            print"},\n";
        }
    }
    print "}\n";
}


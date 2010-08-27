#
# Elaborazione foglio presenze

# Originally developed and customized for A.T.P. srl, Tito (PZ), ITALY
#
# Copyright (C) 2006, 2007 Guido De Rosa <guidoderosa@gmail.com>
#
#    This program is free software; you can redistribute it and/or modify
#    it under the terms of the GNU General Public License as published by
#    the Free Software Foundation; either version 2 of the License, or
#    (at your option) any later version.
#
#    This program is distributed in the hope that it will be useful,
#    but WITHOUT ANY WARRANTY; without even the implied warranty of
#    MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
#    GNU General Public License for more details.
#
#    You should have received a copy of the GNU General Public License
#    along with this program; if not, write to the Free Software
#    Foundation, Inc., 59 Temple Place, Suite 330, Boston, MA  02111-1307  USA


use strict;
#use Data::Dumper;
use Win32::OLE;
use Win32::OLE::Const 'Microsoft Excel';
use Win32::WebBrowser;
use Cwd;
use Tk;
use Tk::Text;
use Tk::Dialog;
use File::Temp qw/ tempfile tempdir /;
use File::Copy;

#$Data::Dumper::Purity = 1;
#$Data::Dumper::Terse = 1;

#constants / config
my $DEBUG = 0;
my $PDFTOTEXT = 'pdftotext.exe';
my $PDFTOTEXTOPTS = '-layout';
my $tmpdir = tempdir( 'Elaborazione_foglio_presenze.XXXX', CLEANUP => 1, TMPDIR => 1 );
my $VERSION = '1.2';
my $APPDATADIR = 'Elaborazione_foglio_presenze';
my $GIORNATA_LAVORATIVA_STD = 8; # ore std = 8
my $maternita_codice_INPS = 'MA';
my $malattia_codice_INPS = 'M';

my $festiviconf = 'festivi.conf';
my $festiviconf_basename = 'festivi.conf';
my $esclusiconf = 'esclusi.conf';
my $esclusiconf_basename = 'esclusi.conf';
my $lastfile_dat = 'lastfile.dat';
my $lastfile_dat_basename = 'lastfile.dat';
configure_files();

my $helpfile = "Documentation/help.html";

#global variables:
my $worker_count = 0; #we're processing the $(worker_count)-th worker; appemd_newrecod() will increment
my $records = []; append_newrecord(); # tha main data structure  # need a constructor? <grin>
my $meseanno = '';
my ($mm, $yy); # month, day, two digits
my $calendar_str = '';
my $days_in_month = 0; # =$last_day anyway... only for compatibility...
my $first_day = 1; # input data may not report from the 1st day of m.
my $last_day = 1; # input data may not report till the very last day
my @festivi = ();
my @calendario_del_mese = ();   # es. undef, 'Me', 'Gi', 'Ve', ... se 
                                # il 1 del mese e` mercoledi`; in italiano 
                                # per restare aderenti al file parsato;
                                # puo' (work needed?) anche essere undef, 
                                # undef, undef, 'Ve', 'Sa', ... se il file 
                                # riporta dati a partire dal 3 del mese
my @esclusi = ();   # $esclusi[codice dip.] = 1 se e' da escludere
my $GUI = 1;

Win32::SetChildShowWindow(0); # make console (child!) window for xpdf invisible 

$records->[$worker_count] = newrecord(); 
    #real data will begin with $records->[1]

my $file = $ARGV[0];
$GUI = not($ARGV[0]); # with cmdline args it's a cmdline tool; else there's a GUI...!
if ($GUI) {GUI()} else {do_all()};

sub reset_globals{
    $worker_count = 0;
    $records = []; append_newrecord(); # tha main data structure  # need a constructor? <grin>
    $meseanno = '';
    $mm=''; $yy='';
    $calendar_str = '';
    $days_in_month = 0;
    $first_day = 1;
    $last_day = 1;
    @festivi = ();
    @calendario_del_mese = ();
    @esclusi = ();
    $records->[$worker_count] = newrecord(); 
}

sub do_all {
    my $textbox = shift if ($GUI); 
    
    parse_esclusi_conf();
    parse_text($file, $textbox);
    calcola_permessi();    
    make_Excel($textbox);
    
}

sub GUI {
    use Tk;    

    my $TOP = MainWindow->new(
        -title      => 'Elaborazione foglio presenze per A.T.P. srl'
    );
    my $toplevel = $TOP->toplevel; 
    #my $toplevel = $TOP;
    my $menubar = $toplevel->Menu(-type => 'menubar');
    $toplevel->configure(-menu => $menubar);
    #$toplevel->geometry('600x210');
    
    my $t = $toplevel->Scrolled(qw/Text -relief sunken -borderwidth 2
			   -height 14 -width 72 -scrollbars oe/);
    #$t->pack(qw/-expand yes -fill both/);
    $t->pack;
    $t->insert('end','Apri un file PDF...');
    #$t->insert('end',"\naaaa");


    my $f = $menubar->cascade(-label => '~File', -tearoff => 0);
    $f->command(-label => 'Apri ...',    -command => sub{file_open_dialog($TOP, $t)} );
    $f->separator;    
    $f->command(-label => 'Esci',        -command => [$TOP => 'destroy']);
    
    my $c = $menubar->cascade(-label => '~Configura', -tearoff => 0);
    $c->command(-label => 'Esclusi', -command => sub{edit_esclusiconf()});
    $c->command(-label => 'Festivi', -command => sub{edit_festiviconf()});

    my $h = $menubar->cascade(-label => '~?', -tearoff => 0);
    $h->command(-label => 'Aiuto',    -command => sub{ help() } );
    $h->command(-label => 'Informazioni sul programma',   -command => sub{about($TOP)} );
    
    MainLoop;
}

sub help {
    open_browser($helpfile);
}

sub about {
    my $TOP = shift;
    my $msg = '
Elaborazione foglio presenze 
Versione '.$VERSION.'

Personalizzato originariamente per 
A.T.P. srl, Tito, PZ, Italia

Copyright (C) 2006, Guido De Rosa 
<guidoderosa@gmail.com> 

This program is distributed under the
GNU General Public License

See Documentation for more details
';
#    $TOP->messageBox(
#        -message => $msg,
#        -title => 'Elaborazione foglio presenze per A.T.P. snc'
#    );
    my $dialog_about = $TOP->DialogBox(
        -title          => 'Elaborazione foglio presenze per A.T.P. snc',
        #-default_button => 'OK',
        -buttons        => ['OK'],
        #-text           => $msg
    );
    $dialog_about->Label(
        -text=>$msg,
        -padx=>16,
        -pady=>8,
        -justify=>'left'
    )->pack;
    
    $dialog_about->Show;
}

sub parse_text {
    my $status = 'BEGIN';
    my $line;
    
    my ($file,$textbox) = @_;
    
    # tmpfile for pdftotext's output
    my ($tmpfh, $tmpfilename) = tempfile( DIR => $tmpdir, SUFFIX => '.txt' );
    
    #open(FILE, "$PDFTOTEXT $PDFTOTEXTOPTS \"$file\" - |");
    system("\"$PDFTOTEXT\" $PDFTOTEXTOPTS \"$file\" \"$tmpfilename\"");
               
    if ($GUI) {
        $textbox->insert('end',"\n\nEstraggo le informazioni dal file PDF:\n");
        $textbox->yview('-pickplace','end');
        $textbox->update;
    }
    
    while(<$tmpfh>) {
        chomp;
        $line = $_;
        $status = status($line, $status, $textbox);        
        #$status = status($line, $status);        
    }
    #close(FILE);    
}

sub status {
    #a first line-parsing...
    # we're working on (pdftotext FOR WINDOWS)'s output - in default 
    # config: no rcfile ;
    # pdftotext version we refer to is part of Xpdf 3.01 ;
    # for Windows, pdftotext and other non-graphical utilities are 
    # little stand-alone (no libaries needed) DOS executables.
    # 
    # WARNING: no other versions! pdftotext from xpdf 3.00 Linux, for example,
    # make a DIFFERENT textfile!!! that this program can't handle!
    
    my $line = shift;
    my $status = shift;
    my $textbox = shift;     
    
    # clean from garbage....
    $line =~ s/Totale Assenze.*$//;
    $line =~ s/OL ore lavorate.*$//;
    $line =~ s/PR permesso retribui.*$//;
    $line =~ s/GG\. lavorati.*$//;
    $line =~ s/OS ore standard.*$//;
    $line =~ s/Eccedenti totali.*$//;
    
    #in the same order we expect to find lines... but also in several 
    #formatting options...
    if ($line =~ m/^\s{0,3}(\S.*\S)\s{3,}C o d i c e Ind\.\s+: (\S+)\s+MESE\/ANNO : (\S+ \S+)\s*$/ or
        $line =~ m/^\s{0,3}(\S.*\S)\s{3,}Matricola\s+: (\S+)\s+MESE\/ANNO : (\S+ \S+)\s*$/ ) { 
        print "\nNOME = '$1'";
        print " COD = '$2'";
        
        if ($GUI) {
            $textbox->insert('end',"\nNome = $1; tessera num. $2");
            $textbox->yview('-pickplace','end');
            $textbox->update;
        }

        $records->[$worker_count]->{'nome'} and append_newrecord();
        $records->[$worker_count]->{'nome'} = $1;
        $records->[$worker_count]->{'codice_dipendente'} and append_newrecord(); # redundant...
        $records->[$worker_count]->{'codice_dipendente'} = $2;
        unless ($meseanno) {
            $meseanno = $3;
        }
        return 'NOME+COD+MESEANNO';
    } 
    if ($line =~ m/Codice Ind\.\s+: (\S+)\s*$/ or $line =~ m/Matricola\s+: (\S+)\s*$/ ) {
        if ($records->[$worker_count]->{'codice_dipendente'}) {
            append_newrecord();
            textboxprint ("\n", $textbox);
        }
        $records->[$worker_count]->{'codice_dipendente'} = $1;
        textboxprint (" Tessera n. $1 ", $textbox);
        return 'CODICE_DIPENDENTE';
    }
    if ( $line =~ m/^(.*\S)\s+MESE\/ANNO : (\S+ \S+)[\s\r]*$/i ) {
        if ($records->[$worker_count]->{'nome'}) {
            append_newrecord();
            textboxprint ("\n", $textbox);
        }
        $records->[$worker_count]->{'nome'} = $1;
        unless ($meseanno) {
            $meseanno = $2;
        }
        textboxprint (" Nome = '$1' ", $textbox);
        return 'NOME_MESE_ANNO';
    }
    if ( $line =~ m/(([a-z]{2}[0-9]{2} )+([a-z]{2}[0-9]{2}))\s*$/i ) {

        unless ($calendar_str) {
            $calendar_str = $1;
            month();
        }
        # A questo punto assumiamo di avere il calendario del mese e i dati
        # del dipendente; dobbiamo riempire i suoi 'results' per tutto il
        # mese con i dati di default: $GIORNATA_LAVORATIVA_STD ore lavorate, 0 straord. e 0 permessi
        # (addendum: 0 ore di "maternita' ore": anche se non verra' stampata, bisogna 
        # tenerne memoria per calcola_permessi)
        # per i giorni feriali; no straord. festvi (SE IL DIPENDENTE NON HA
        # A MATRICOLA O CODICE => E' STATO ASSUNTO DURANTE IL MESE (O PORZIONE DI 
        # MESE...) E LO SI ESClUDE DAL CALCOLO PER EVITARE DI ATTRIBUIRGLI ORE NON
        # LAVORATE: COMPARIRà UNA RIGA BIANCA...)
        if ( has_code($worker_count) ) {
            set_default_worker_results();
        }
        # TODO: else? set_results_to('_') ? 
        return 'CALENDAR';  # of the month
    }
    return 'IGNORED' unless has_code($worker_count);
    return 'IGNORED' if ($records->[$worker_count]->{'escludi'});
    if ($line =~ m/^\s{0,2}STR\s+(\S.+\S)[\s\r]*$/) { # straordinario feriale
        straordinario_feriale($1);
        return 'STR';
    }
    if ($line =~ m/^\s{0,2}STRF\s+(\S.+\S)[\s\r]*$/) { # straordinario fest
        straordinario_festivo($1);
        return 'STRF';
    }
    if ($line =~ m/^\s{0,2}FER\s+(\S.+\S)[\s\r]*$/) { # ferie (in giorni)
        ferie($1);
        return 'FER';
    }
    if ($line =~ m/^\s{0,2}PR\s+(\S.+\S)[\s\r]*$/) { # permesso retribuito
        permesso_retribuito($1);
        return 'PR';
    }
    if ($line =~ m/^\s{0,2}MAL\s+(\S.+\S)[\s\r]*$/) { # malattia (in giorni)
        malattia($1);
        return 'MAL';
    }
    if ($line =~ m/^\s{0,2}MATO\s+(\S.+\S)[\s\r]*$/) { # maternità ore 
        maternita_ore($1);
        return 'MATO';
    }
    if ($line =~ m/^\s{0,2}MATG\s+(\S.+\S)[\s\r]*$/) { # maternità giorni
        maternita_giorni($1);
        return 'MATG';
    }
    return 'OTHER';
}

sub newrecord {
    my $record =  {
        'nome'              => '',  # hum. read. lastname+firstname
        'codice_dipendente' => '',  # num. di tessera     
        'results'           => [ 0 => undef ],   # risultati per ogni giorno del mese
	    'escludi'           => 0
    };

    return $record;
}

sub append_newrecord {
    $worker_count++;
    $records->[$worker_count] = newrecord();    	
}

sub month { # imposta il calendario del mese e i gg festivi
    my ($gsett, $gmese, $gsettgmese);
    my $str = $calendar_str;
    $str =~ m/^\s*[a-z]{2}([0-9]{2}).*[a-z]{2}([0-9]{2})[\r\s]*$/i;
    $first_day = $1; # if input file doesn't report from the 1st
    $last_day = $2; # last day in month
    $days_in_month = $last_day; # for backward compatibily... anything todo?

    my @ary = split(/\s+/, $str);
    $festivi[0] = undef; $calendario_del_mese[0] = undef;
    foreach $gsettgmese(@ary) {
        $gsettgmese =~ m/([a-z]{2})([0-9]{2})/i;
        $gsett = $1; $gmese = $2; 
        $calendario_del_mese[$gmese] = "$gsett"; 
        if ( ($gsett eq 'Sa') or ($gsett eq 'Do') ) {
            $festivi[$gmese] = 1;
        } else { $festivi[$gmese] = 0; }
    }
    
    $meseanno =~ m/([a-z]+)[^a-z0-9]+\d{0,2}(\d{2})/i;
    $yy = $2; # 'return'
    $mm = mese2digits($1);    
    
    aggiungi_festivi_da_file();
}

sub mese2digits {
    my $str = shift;
    $str =~ m/gennaio/i and return '01';
    $str =~ m/febbraio/i and return '02';
    $str =~ m/marzo/i and return '03';
    $str =~ m/aprile/i and return '04';
    $str =~ m/maggio/i and return '05';
    $str =~ m/giugno/i and return '06';
    $str =~ m/luglio/i and return '07';
    $str =~ m/agosto/i and return '08';
    $str =~ m/settembre/i and return '09';
    $str =~ m/ottobre/i and return '10';
    $str =~ m/novembre/i and return '11';
    $str =~ m/dicembre/i and return '12';
}

sub aggiungi_festivi_da_file {
    open(FESTIVICONF,"<$festiviconf") or return;
    while(<FESTIVICONF>) {
        chomp;
        s/#.*$//; # ignore comments
        if (m/^\s*(\d{2})(\d{2})[^\d]*$/) { ## mmdd
            if ($1 == $mm) {
                $festivi[$2] = 1;
            }
        }
        if (m/^\s*(\d{2})(\d{2})(\d{2})[^\d]*$/) { ## yymmdd
            if (($1 == $yy) and ($2 == $mm)) {
                $festivi[$3] = 1;
            }
        }
    }
    close(FESTIVICONF);
}

sub parse_esclusi_conf {
    open(ESCLUSI,"<$esclusiconf") or return;
    while(<ESCLUSI>){
        chomp;
        s/#.*$//; # ignore comments
        m/^[^\d]*(\d+)[^\d]*$/ and ($esclusi[$1] = 1);
    }
    close(ESCLUSI);
}



sub straordinario_feriale {
    my $str = shift;
    my $day;


    my @ary = split(/\s+/,$str);

    for ($day=$first_day;$day<=$last_day;$day++) {
        $records->[$worker_count]->{'results'}->[$day]->{'O'} 
        += ore($ary[$day-$first_day]) unless ($festivi[$day]);
    }
}
sub straordinario_festivo {
    my $str = shift;
    my $day;


    my @ary = split(/\s+/,$str);

    for ($day=$first_day;$day<=$last_day;$day++) {
        $records->[$worker_count]->{'results'}->[$day]->{'Fes'} 
        = ore($ary[$day-$first_day]) if ($festivi[$day]);
    }
}
sub ferie {
    my $str = shift;
    my $day;


    my @ary = split(/\s+/,$str);
    for ($day=$first_day;$day<=$last_day;$day++) {
        $records->[$worker_count]->{'results'}->[$day]->{'O'}
        = 'F' if $ary[$day-$first_day] =~ m/^[a-z0-9]+$/i; # tipicam '0100'
    }
}
sub permesso_retribuito {
# questa sub non imposta il valore di 'Per', si limita a scalare le ore lavorate;
# 'Per' sara' calcolato quando si avranno tutti i dati: se ad esempio ci sono 2 ore di
# permesso e due ore di str lo stesso giorno, il risultato dev'essere 
# O=$GIORNATA_LAVORATIVA_STD, Per=0 e non O=$GIORNATA_LAVORATIVA_STD, Per=2
#
    my $str = shift;
    my $day;


    my @ary = split(/\s+/,$str);
    for ($day=$first_day;$day<=$last_day;$day++) {
        if ($records->[$worker_count]->{'results'}->[$day]->{'O'} =~ 
        m/^[\s\.0-9]+$/) { # is numeric...
            $records->[$worker_count]->{'results'}->[$day]->{'O'}
            -= ore($ary[$day-$first_day]);
            # $records->[$worker_count]->{'results'}->[$day]->{'Per'}
            # += ore($ary[$day-$first_day]); # NO! will be calculated later... calcola_permessi()
        }
    }
}
sub malattia {
    my $str = shift;
    my $day;    


    my @ary = split(/\s+/,$str);
    for ($day=$first_day;$day<=$last_day;$day++) {
        $records->[$worker_count]->{'results'}->[$day]->{'O'}
        = $malattia_codice_INPS if $ary[$day-$first_day] =~ m/^[a-z0-9]+$/i; 
    }
}
sub maternita_ore {
    my $str = shift;
    my $day;

    my @ary = split(/\s+/,$str);
    for ($day=$first_day;$day<=$last_day;$day++) {

        if ($records->[$worker_count]->{'results'}->[$day]->{'O'} =~ 
        m/^[\s\.0-9]+$/) { # is numeric...
            
            $records->[$worker_count]->{'results'}->[$day]->{'O'}
            -= ore($ary[$day-$first_day]);
        }
    }
}
sub maternita_giorni {
    my $str = shift;
    my $day;    


    my @ary = split(/\s+/,$str);
    for ($day=$first_day;$day<=$last_day;$day++) {
        $records->[$worker_count]->{'results'}->[$day]->{'O'}
        = $maternita_codice_INPS if $ary[$day-$first_day] =~ m/^[a-z0-9]+$/i; 
    }
}

sub set_default_worker_results {    
    my $i;
    
    if ( $esclusi[ $records->[$worker_count]->{'codice_dipendente'} ] ) {
        $records->[$worker_count]->{'escludi'} = 1;
        return;
    }

    for ($i=$first_day;$i<=$last_day;$i++) {        
        if ($festivi[$i]) {
            $records->[$worker_count]->{'results'}->[$i] = {
                'O' => '_fest', 'Fes' => 0, 'Per' => '_fest'    
            };
        } else { #feriale 
            $records->[$worker_count]->{'results'}->[$i] = {
                'O' => $GIORNATA_LAVORATIVA_STD, 'Fes' => '', 'Per' => 0, 'MATO' => 0
            }   # 'O'=ore lavorate, 'Fes'=straord. festivo, 'Per'=permesso retriuito,
                # 'MATO' = maternita' ore (non verra' stampata ma bisogna tenerne memoria
                # per calcola_permessi() )
                # '_fest'=non ci sono ore lavorate regolari nei gg festivi...
        }
    }

} 

sub ore { 
    my $str = shift;
    ($str =~ m/^\s*([0-9]{1,2})([0-9]{2})\s*$/ ) or return 0; 
    return ($1 + ($2/60));   # TODO? rounding?
}

sub has_code {
    my $n;
    ($n = shift) or $n = $worker_count; # has_code with or without argments..
    
    $records->[$n]->{'codice_dipendente'} =~ m/^[^a-z0-9]*$/i and return 0;
    return $records->[$n]->{'codice_dipendente'};
}

sub calcola_permessi { # 'a posteriori'...
    my ($day, $worker, $MATO, $O) = (1, 1, 0, "");
    
    for $worker (1..$worker_count) {
        if ( (!$records->[$worker]->{'escludi'}) and has_code($worker) ) {
            for $day ($first_day..$last_day) {
                unless ($festivi[$day]) {                    
                    $MATO   = $records->[$worker]->{'results'}->[$day]->{'MATO'};
                    $O      = $records->[$worker]->{'results'}->[$day]->{'O'};
                    #print "\n".$records->[$worker]->{'nome'}." \$O=$O \$worker=$worker \$day=$day \$MATO=$MATO\n" unless $O; # DEBUG
            
                    if ( ( $O =~ m/^[\s\d\.\,]+$/ ) and ( ($GIORNATA_LAVORATIVA_STD - $MATO) > $O ) ) {
                        $records->[$worker]->{'results'}->[$day]->{'Per'} = $GIORNATA_LAVORATIVA_STD - $MATO - $O;    
                    }
                }
            }
        }
    }
}

sub make_Excel {    
    my $textbox = shift;
    
    print "\nGenerando il foglio Excel... Attendere, prego...";    
    
    if ($GUI) {
        $textbox->insert('end',"\n
Creazione del foglio Excel... 
        
A seconda della velocità del computer
questa operazione potrebbe richiedere 
alcuni secondi...\n");
        $textbox->yview('-pickplace','end');
        $textbox->update;
    }
	
    my ($day, $worker, $Range, $Cell1, $Cell2, $Range_row, $Cell3, $row); 
    my $Excel = Win32::OLE->new("Excel.Application");
    my $Book = $Excel->Workbooks->Add(cwd.'\atp.xlt');
    my $Sheet = $Book->Worksheets(1);
    
    #$Excel->{Visible} = 1; # DEBUG?
    
    $Sheet->Cells(1,3)->{Value} = " ".$meseanno;

    if ($GUI) {
        $textbox->insert('end',"\nGenerazione calendario...");
        $textbox->yview('-pickplace','end');
        $textbox->update;
    }
    
    for $day ($first_day..$last_day) {
        print ".";
        if ($GUI) {
            $textbox->insert('end','.'); # a "progress" point
            $textbox->yview('-pickplace','end');
            $textbox->update;
        }
        $Sheet->Cells(3,3+$day)->{Value} = $day;
        $Sheet->Cells(4,3+$day)->{Value} = $calendario_del_mese[$day];
        $Range = $Sheet->Range($Sheet->Cells(4,3+$day), $Sheet->Cells(4 + 3*$worker_count, 3+$day));
        $Range->Interior->{Color} = 0x7799FF if ($festivi[$day]); #pink..?
    }
    
    if ($GUI) {
        $textbox->insert('end',"\n\nScrittura report...");
        $textbox->yview('-pickplace','end');
        $textbox->update;
    }
    
    for $worker (1..$worker_count) {
        print ".";
        if ($GUI) {
            $textbox->insert('end','.') ; # a "progress" point
            $textbox->yview('-pickplace','end');
            $textbox->update;
        }
        $row = 2 + 3*$worker; 
        $Cell1 = $Sheet->Cells($row, 1);
        $Cell2 = $Sheet->Cells($row + 2, 2);
        $Cell3 = $Sheet->Cells($row + 2, 36);
        $Range = $Sheet->Range($Cell1, $Cell2);
        $Range->Merge();
        $Range_row = $Sheet->Range($Cell1, $Cell3);
        if ($records->[$worker]->{'escludi'} or ($records->[$worker]->{'codice_dipendente'} =~ m/^[^a-z0-9]*$/i) ) {
            $Cell1->{Value} = "$records->[$worker]->{'nome'} *\ntessera n. $records->[$worker]->{'codice_dipendente'}";
            $Cell1->{Font}->{Color} = 0x0000ff;
        } else {        
            $Cell1->{Value} = "$records->[$worker]->{'nome'}\ntessera n. $records->[$worker]->{'codice_dipendente'}";
        }        
        $Range_row->Borders->{Weight} = xlThin;
        $Range_row->Borders(xlEdgeTop)->{Weight} = xlMedium;
        $Sheet->Cells($row, 3)->{Value} = 'O';
        $Sheet->Cells($row + 1, 3)->{Value} = 'Fes';
        $Sheet->Cells($row + 2, 3)->{Value} = 'Per';
        
        for $day ($first_day..$last_day) {
            $Sheet->Cells($row   , 3 + $day)->{Value} = $records->[$worker]->{'results'}->[$day]->{'O'} if
            ( $records->[$worker]->{'results'}->[$day]->{'O'} and ($records->[$worker]->{'results'}->[$day]->{'O'} ne '_fest') );
            if ( $records->[$worker]->{'results'}->[$day]->{'Fes'} ) {
                $Sheet->Cells($row +1, 3 + $day)->{Value} = $records->[$worker]->{'results'}->[$day]->{'Fes'};
                $Sheet->Cells($row +1, 3 + $day)->Interior->{Color} = 0xFFFFFF;
            }
            $Sheet->Cells($row +2, 3 + $day)->{Value} = $records->[$worker]->{'results'}->[$day]->{'Per'} if
            ( $records->[$worker]->{'results'}->[$day]->{'Per'} and ($records->[$worker]->{'results'}->[$day]->{'Per'} =~ m/[1-9]+/) ); 

        }
        
        $Sheet->Cells($row    , 35)->{Formula} = "=SOMMA(D" . ($row)   . ":AH" . ($row)   . ")" ;
        $Sheet->Cells($row + 1, 35)->{Formula} = "=SOMMA(D" . ($row+1) . ":AH" . ($row+1) . ")" ;
        $Sheet->Cells($row + 2, 35)->{Formula} = "=SOMMA(D" . ($row+2) . ":AH" . ($row+2) . ")" ;
        $Sheet->Cells($row    , 36)->{Formula} = "=SOMMA(AI". ($row)   . ":AI" . ($row+1) . ")" ; 
        
        $Cell1 = $Sheet->Cells($row, 36);
        $Cell2 = $Sheet->Cells($row + 1, 36);
        $Range = $Sheet->Range($Cell1, $Cell2);
        $Range->Merge();

    }
        
    $Sheet->Range( $Sheet->Cells(3,35), $Sheet->Cells(4 + 3*$worker_count, 35) )->Borders(xlEdgeLeft)->{Weight}  = xlMedium;
    $Sheet->Range( $Sheet->Cells(5,3),  $Sheet->Cells(4 + 3*$worker_count,  3) )->Borders(xlEdgeLeft)->{Weight}  = xlMedium;
    $Sheet->Range( $Sheet->Cells(5,3),  $Sheet->Cells(4 + 3*$worker_count,  3) )->Borders(xlEdgeRight)->{Weight} = xlMedium;
    
    print "OK\n";
    if ($GUI) {
        $textbox->insert('end',"Fatto.\n\n");
        $textbox->yview('-pickplace','end');
        $textbox->update;        
    }
    
    $Excel->{Visible} = 1;
}

sub file_open_dialog {
    my $top = shift;
    my $textbox = shift;
    
    reset_globals();
    
    # my $lastfile_dat='lastfile.dat'; # now it's global/configurable!
    open(LASTFILEDAT,"<$lastfile_dat");
    my $lastfile = <LASTFILEDAT>;
    close(LASTFILEDAT);
    $lastfile =~ s/\//\\/g;
    my $lastdir=$lastfile;
    $lastdir =~ s/^(.*[\\\/])[^\\\/]+$/$1/;


    my $types = [
#     ['Text Files',       ['.txt', '.text']],
#     ['TCL Scripts',      '.tcl'           ],
#     ['C Source Files',   '.c',      'TEXT'],
#     ['GIF Files',        '.gif',          ],
      ['PDF Files',        '.pdf'           ],
      ['All Files',        '*'            ]
    ];


    $file = $top->getOpenFile(
        -initialdir => $lastdir,
        -filetypes=>$types
    );

    if ($file) {
        open(FILE,">$lastfile_dat");
        print FILE $file; 
        close(FILE);
        $file =~ s/\//\\/g;
        do_all($textbox);
        #do_all();
    }
}

sub textboxprint {
    my $txt = shift;
    my $textbox = shift;
    
    print $txt;
    unless($GUI) {
            return;
    }    
    $textbox->insert('end',$txt);
    $textbox->yview('-pickplace','end');
    $textbox->update; 
}

sub configure_files {
    return 0 unless $ENV{APPDATA};
    my $APPDATADIRPATH = $ENV{APPDATA}.'/'.$APPDATADIR;
    
    -d $APPDATADIRPATH or mkdir $APPDATADIRPATH;
    -f "$APPDATADIRPATH/$festiviconf_basename" or copy (cwd."/".$festiviconf_basename, "$APPDATADIRPATH/$festiviconf_basename");
    -f "$APPDATADIRPATH/$esclusiconf_basename" or copy (cwd."/".$esclusiconf_basename, "$APPDATADIRPATH/$esclusiconf_basename");
    -f "$APPDATADIRPATH/$lastfile_dat_basename" or copy (cwd."/".$lastfile_dat_basename, "$APPDATADIRPATH/$lastfile_dat_basename");
    $festiviconf = "$APPDATADIRPATH/$festiviconf_basename";
    $esclusiconf = "$APPDATADIRPATH/$esclusiconf_basename";
    $lastfile_dat = "$APPDATADIRPATH/$lastfile_dat_basename";
    
    return 1;
}

sub edit_esclusiconf{
    system($esclusiconf);
}

sub edit_festiviconf{
    system($festiviconf);
}
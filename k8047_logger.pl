# k8047_logger.pl
# Glenn Axelrod 2011-11-27

# Perl program reads from Velleman K8047 4-channel USB data recorder using
# their k8047D.dll

# Log the data using real-time clock timestamps in ISO 8601 format. This program  
# was written to log parameters over days or weeks so it uses the generic 
# Perl one second sleep as the minimum interval. 

# Get the components from Velleman:
# http://www.velleman.be/downloads/files/downloads/k8047_pcs10_dll.zip
# Extract the files and copy them to this location:
# C:\WINDOWS\system32\FASTTime32.dll [1998]
#                     K8047D.dll     [2011-01-27]
#                     K8047E.exe     [2011-06-22]
# Check the creation dates to be sure you have the latest versions.

# Acquire the Win32::API Perl module to access DLL functions. It's at
# http://search.cpan.org/CPAN/authors/id/A/AC/ACALPINI/Win32-API-0.41.tar.gz
# But after installing Strawberry Perl, the command "cpan Win32::API" was 
# sufficient.
# Articles in www.perlmonks.org/?node_id=853418 and 854708 were helpful
# for understanding the use of pack to create a pre-allocated buffer pointer
# to pass to ReadData.

# Useful utilities from the Unix world include "tail -f" which lets you watch
# the logfile updating live in another window. Check MS for Windows Resource Kits.

# Program was written on 32-bit Windows XP SP3. 
# TBD: What happens on 64-bit Windows 7?

use Win32::API qw();
use Getopt::Long qw(:config ignore_case);
use strict;

# The API has a debug function.
 $Win32::API::DEBUG=1;

# Exit handler allows Ctrl/C to shut things down safely.
$SIG{INT} = 'handler';
$SIG{QUIT} = 'handler';

my $interactive = 1; # 1 means Print output to screen with channel labels
                     # 0 means Write to file in csv format
my $sample_interval = 1; # In seconds
my $run_for = 0; # In seconds
my $path_to_logfile = ".";
my $logfile = "default_log.txt";
my $total_samples = 0;
my $run_time = 0;
my ($debug,$help) = (0,0);
my $scale = 0; # If a string of values, allow converting [0-255] 
               # readings to voltages per list
my @rangelist = (); # Set Full Scale range values for each of four channels
my @gainlist = (1,1,1,1); # Default gain settings
my ($chan1,$chan2,$chan3,$chan4) = (0,0,0,0);

# Constants for scale conversion from 255 counts to full scale voltage 
# (fs_volts) for display 
my %fs_volts = (
   "30" => eval(30/255),
   "15" => eval(15/255),
   "6"  => eval(6/255),
   "3"  => eval(3/255)
);

# Constants for setting the input gain on the data recorder
my %gain = (
   "30" => 1,
   "15" => 2,
   "6"  => 5,
   "3"  => 10
);

my %getopts = (
   "h|help"              => \$help,
   "debug"               => \$debug,

   "i|interactive!"      => \$interactive,
   "s|sample_interval:i" => \$sample_interval,
   "r|run_for:i"         => \$run_for,
   "l|logfile:s"         => \$logfile,
   "p|path_to_logfile:s" => \$path_to_logfile,
   "scale:s"             => \$scale
   );
   
     
my $help_text = "
     Usage: perl k8047.pl  -option <val>
     \nk8047.pl options: \n
     -help | -interactive|noi -sample_interval=seconds
     -run_for=seconds (0=no stop) -logfile=name -path_to_file=.
     -scale=3,6,15,30 (volts, one entry for each channel)
     (10mV sensitivity, +Vdc only) -debug
     perl  k8047.pl -i -s=60 -r=86400 -l=each_minute_for_one_day.log
     ";

GetOptions(%getopts) || die "*** Can't parse options\n" . $help_text;

if ($help) {
    print $help_text;
    exit 0;
}

my $o_file = "$path_to_logfile/$logfile";

print "Interactive: $interactive\n", 
"Sample Interval: $sample_interval\n", 
"Run For: $run_for\n", 
"Log File: $o_file\n",
"Scale = $scale\n";

if ($debug) {
   if ($scale) {
      print "Full scale volts conversion and gain table:\n",
         "  30 = $fs_volts{'30'}, gain = $gain{'30'}\n",
         "  15 = $fs_volts{'15'}, gain = $gain{'15'}\n",
         "   6 = $fs_volts{'6'}, gain = $gain{'6'}\n",
         "   3 = $fs_volts{'3'}, gain = $gain{'3'}\n\n";
   } else {
      print "Scale is $scale; all channels default 30 volts\n";
   }
}


if ($scale) {
   @rangelist = split(",", $scale);
   print "Full scale voltage ranges for channels 1-4 = @rangelist\n" if $debug;
   
   # Convert ranges (3,6,15,30) to gains (10,5,2,1) to set each channel.
   for (my $channel=1; $channel<5; $channel++) {
       # @gainlist is 0-based so try to keep track of channel number
       my $ch_index = eval($channel - 1);
       $gainlist[$ch_index] = $gain{$rangelist[$ch_index]};
   }
    print "Channel gain list = @gainlist\n" if $debug;
}

# Format of the return values from the Velleman K8047
# Timestamps are since the device started up; not using them.
# The last two values are reserved by the DLL.
my ($timer_lsb,$timer_msb,$ch1,$ch2,$ch3,$ch4,$res1,$res2) = 
   (0,0,0,0,0,0,0,0);
my @data = ($timer_lsb,$timer_msb,$ch1,$ch2,$ch3,$ch4,$res1,$res2);

my $import = Win32::API->Import("K8047D", "StartDevice", "", "");

# Report results of the first DLL function import attempt. If it 
# succeeds, we probably have everything set up correctly. Otherwise, 
# there's a problem so we should stop. Import returns 1 for success.

if ($import) {
   print "Import StartDevice success.\n";
} else
{
   print "Import results: $import\nImport message is: $^E\n";
   exit 1;
}

Win32::API->Import("K8047D", "StopDevice", "", "");
Win32::API->Import("K8047D", "ReadData", "P", "I");
Win32::API->Import("K8047D", "LEDon", "", "");
Win32::API->Import("K8047D", "LEDoff", "", "");
Win32::API->Import("K8047D", "SetGain", "II", "I");
Win32::API->Import("K8047D", "Connected", "", "I");

print "Done importing procedures\n" if $debug;

my $start = StartDevice();

# It doesn't seem to harm anything, but when StartDevice() runs, 
# it causes a windows popup to flash/blink on and then disappear
# too fast to see. Problem is fixed in USB Utility K8047E.exe 2011-06-22.
# http://forum.velleman.eu/viewtopic.php?f=10&t=6221&hilit=frequent+window

# Verify there is a device connected by USB. Otherwise, the commands will 
# run without complaint without any device there, returning 0 data. The 
# Connected command is a feature in K8047D.DLL from 2011-01-27. I think 
# it may work by implementing a suggestion in the forum to verify that 
# the time stamp from ReadData is incrementing over multiple reads.
# The device has to be started first. Refer to the forum for details:
# http://forum.velleman.eu/
#   viewtopic.php?f=10&t=5660&hilit=k8047+connected+reset+detection

my $device_connected = Connected();
print "Device connected return value: $device_connected\n";

if (!$device_connected) {
   # Override the test to allow debugging the program without
   # the device present.
   if ($debug) {
      print "No device. Debug mode for program only.\n";
   } else {
      print "No device found on USB port! Exiting.\n";
      exit 0;
   }
}


if ($scale) {
   # Turn on the Record light during Gain setting
   my $lon_st = LEDon();
   print "LEDon status = $lon_st\n" if $debug;
   print "Setting gains\n" if $debug;
   sleep 1;

   for (my $channel=1; $channel<5; $channel++) {
      # @gainlist array is 0-based so try to keep track of channel number
      my $ch_index = eval($channel - 1);
      print "Ch $channel gain $gainlist[$ch_index]\n";
      my $gain_status = SetGain($channel, $gainlist[$ch_index]);
      print "SetGain Ch $channel status = $gain_status\n" if $debug;
   }
   my $lof_st = LEDoff();
   print "LEDoff status = $lof_st\n" if $debug;
}

my $loop_index = 0;

if ($run_for == 0) {
   $loop_index = 1;
} else {
   $loop_index = $run_for;
}

print "Loop index = $loop_index\n" if $debug;

while ($loop_index>0) {
   # Set up a buffer to hold 8 integers
   my $r_data = pack('llllllll', 0,0,0,0,0,0,0,0);
   my $lon_st = LEDon();
   my $rstatus = ReadData($r_data);
   print "Read status = $rstatus\n" if $debug;
   my $lof_st = LEDoff();
   @data = unpack('llllllll', $r_data);

   my ($sec,$min,$hour,$mday,$mon,$year,$wday,$yday,$isdst)
       = gmtime(time);
   $year = $year + 1900;
   $mon = $mon + 1; # 0-based
   
   # Zulu (GMT) time in one-part ISO 8601 format makes sorting on dates easy
   my $timestamp = sprintf("%4d-%02d-%02dT%02d:%02d:%02dZ", 
                         $year,$mon,$mday,$hour,$min,$sec);
   
   if ($scale) {
      # Values corrected to full scale as defined per channel
      $chan1 = sprintf("%2.2f", eval($data[2] * $fs_volts{$rangelist[0]}));
      $chan2 = sprintf("%2.2f", eval($data[3] * $fs_volts{$rangelist[1]}));
      $chan3 = sprintf("%2.2f", eval($data[4] * $fs_volts{$rangelist[2]}));
      $chan4 = sprintf("%2.2f", eval($data[5] * $fs_volts{$rangelist[3]}));
   } else {
      # Raw 0..255 values
      $chan1 = $data[2];
      $chan2 = $data[3];
      $chan3 = $data[4];
      $chan4 = $data[5]; 
   }
   
   if ($interactive) {
      # Spell out channel numbers
      my $datalist = "Ch1:$chan1, Ch2:$chan2, Ch3:$chan3, Ch4:$chan4";
      print "$timestamp $datalist\n";
   } else {
      # Output unlabelled csv format for easy import into spreadsheet
      my $datalist = "$chan1,$chan2,$chan3,$chan4";

      if ($debug) {
         # Write to console
         print "debug:$timestamp,$datalist\n";
      } else {
         # Write to file, opening and closing on each write. 
         # Makes tail, copy, etc easy without affecting the log session.
         open OUTPUT, ">>$o_file";
         print OUTPUT "$timestamp,$datalist\n";
         close OUTPUT;
      }
   }

   # Increment counter for end report
   $total_samples++;
   
   if ($run_for > 0) {
      $loop_index = $loop_index - $sample_interval;
   }
   print "Loop index after decrement = $loop_index\n" if $debug;
   
   # Perl on Windows doesn't let Ctrl/C interrupt long sleeps. For 
   # long sample intervals, this makes the program hard to stop. So, 
   # implement the interval as a series of short sleeps.
   for (my $sl_index=0; $sl_index<$sample_interval; $sl_index++) { 
      sleep 1;
   }  
}
   my $lof_st = LEDoff();
   print "LEDoff status = $lof_st\n";
   handler();

sub handler {
   my ($sig) = @_;
   my $lof_st = LEDoff();
   my $end = StopDevice();
   close OUTPUT;

   if ($interactive) {
       # $run_for might be 0 so calculate actual elapsed time
       my $runtime_seconds = $total_samples * $sample_interval;
       my $runtime = dhms($runtime_seconds); # Convert to DD:HH:MM:SS
       print "Total samples:  $total_samples; ",
             "Interval: $sample_interval, Run time = $runtime\n";
   }
    exit 0;
}

# Convert run-time from seconds to days, etc
sub dhms {
   # 90061 seconds is 1:1:1:1 
   my ($seconds) = @_;
   my $days = int($seconds/86400);
   my $remainder = $seconds - (eval($days*86400));
   my $hours = int($remainder/3600);
   $remainder = $remainder - (eval($hours*3600));
   my $minutes = int($remainder/60);
   $seconds = $remainder - (eval($minutes*60));
   return "$days Days, $hours Hours, $minutes Min, $seconds Secs";
}
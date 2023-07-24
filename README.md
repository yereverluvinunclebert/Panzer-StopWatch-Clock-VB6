# Panzer-StopWatch-Clock-VB6

 A FOSS Stopwatch VB6 Widget for Reactos, XP, Win7, 8 and 10+.
 
This Panzer widget is an attractive dieselpunk Yahoo widget for your desktop. 
It will be a simple multi-timezone stopwatch and clock. Functional and gorgeous at 
the same time. This Widget is a moveable widget that you can move anywhere 
around the desktop as you require.

![panzer-stopwatch650](https://github.com/yereverluvinunclebert/Panzer-Clock-Ywidget/assets/2788342/352fb2cb-742d-4e24-afe7-906b05e7fb1e)

This widgets functionality is limited as it is just a template for widgets yet
to come, however, it can be increased in size, animation speed can be changed, 
opacity/transparency may be set as to the users discretion. The widget can 
also be made to hide for a pre-determined period.

![tank-clock-mk1](https://github.com/yereverluvinunclebert/Panzer-StopWatch-Clock-VB6/assets/2788342/45805383-244f-4370-ba3e-3259b9fd3805)

Right-click on the widget to display the function menu, mouse hover over the 
widget and press CTRL+mousewheel up/down to resize. It works well on Windows XP 
to Windows 11.

![panzer-clock-web-help](https://github.com/yereverluvinunclebert/Panzer-StopWatch-Clock-VB6/assets/2788342/62704796-76ee-4053-a163-c0767d2cd42b)

The Panzer Stopwatch VB6 gauge is Beta-grade software, under development, not yet 
ready to use on a production system - use at your own risk.

This version was developed on Windows 7 using 32 bit VisualBasic 6 as a FOSS 
project creating a WoW64 widget for the desktop. 

It is open source to allow easy configuration, bug-fixing, enhancement and 
community contribution towards free-and-useful VB6 utilities that can be created
by anyone. The first step was the creation of this template program to form the 
basis for the conversion of other desktop utilities or widgets. A future step 
is new VB6 widgets with more functionality and then hopefully, conversion of 
each to RADBasic/TwinBasic for future-proofing and 64bit-ness. 

This utility is one of a set of steampunk and dieselpunk widgets. That you can 
find here on Deviantart: https://www.deviantart.com/yereverluvinuncleber/gallery

I do hope you enjoy using this utility and others. Your own software 
enhancements and contributions will be gratefully received if you choose to 
contribute.

BUILD: The program runs without any Microsoft plugins.

Built using: VB6, MZ-TOOLS 3.0, VBAdvance, CodeHelp Core IDE Extender
Framework 2.2 & Rubberduck 2.4.1, RichClient6

Links:

	https://www.vbrichclient.com/#/en/About/
	MZ-TOOLS https://www.mztools.com/  
	CodeHelp http://www.planetsourcecode.com/vb/scripts/ShowCode.asp?txtCodeId=62468&lngWId=1  
	Rubberduck http://rubberduckvba.com/  
	Rocketdock https://punklabs.com/  
	Registry code ALLAPI.COM  
	La Volpe http://www.planet-source-code.com/vb/scripts/ShowCode.asp?txtCodeId=67466&lngWId=1  
	PrivateExtractIcons code http://www.activevb.de/rubriken/  
	Persistent debug code http://www.vbforums.com/member.php?234143-Elroy  
	Open File common dialog code without dependent OCX - http://forums.codeguru.com/member.php?92278-rxbagain  
	VBAdvance  
	Fafalone for the enumerate Explorer windows code


Tested on :

	ReactOS 0.4.14 32bit on virtualBox    
	Windows 7 Professional 32bit on Intel    
	Windows 7 Ultimate 64bit on Intel    
	Windows 7 Professional 64bit on Intel    
	Windows XP SP3 32bit on Intel    
	Windows 10 Home 64bit on Intel    
	Windows 10 Home 64bit on AMD    
	Windows 11 64bit on Intel  
	
CREDITS:

I have really tried to maintain the credits as the project has progressed. If I 
have made a mistake and left someone out then do forgive me. I will make amends 
if anyone points out my mistake in leaving someone out.

MicroSoft in the 90s - MS built good, lean and useful tools in the late 90s and 
early 2000s. Thanks for VB6.

Olaf Schmidt - This tool was built using the RichClient RC5 Cairo wrapper for 
VB6. Specifically the components using transparency and reading images directly 
from PSD. Thanks for the massive effort Olaf in creating Cairo counterparts for 
all VB6 native controls and giving us access to advanced features on controls 
such as transparency.

Shuja Ali @ codeguru for his settings.ini code.

ALLAPI.COM        For the registry reading code.

Rxbagain on codeguru for his Open File common dialog code without a dependent 
OCX - http://forums.codeguru.com/member.php?92278-rxbagain

si_the_geek       for his special folder code

Elroy on VB forums for the balloon tooltips

Harry Whitfield for his quality testing, brain stimulation and being an 
unwitting source of inspiration. 

Dependencies:

o A windows-alike o/s such as Windows XP, 7-11 or Apple Mac OSX 11. 

o Microsoft VB6 IDE installed with its runtime components. The program runs 
without any additional Microsoft OCX components, just the basic controls that 
ship with VB6.  

	
* Uses the RC6 Cairo wrapper from Olaf Schmidt.

During development the RC6 components need to be registered. These scripts are 
used to register. Run each by double-clicking on them.

RegisterRC6inPlace.vbs
RegisterRC6WidgetsInPlace.vbs

During runtime on the users system, the RC6 components are dynamically 
referenced using modRC6regfree.bas which is compiled into the binary.	

Requires a PzStopWatch folder in C:\Users\<user>\AppData\Roaming\ 
eg: C:\Users\<user>\AppData\Roaming\PzStopWatch
Requires a settings.ini file to exist in C:\Users\<user>\AppData\Roaming\PzStopWatch
The above will be created automatically by the compiled program when run for the 
first time.

Uses just one OCX control extracted from Krools mega pack (slider). This is part 
of Krools replacement for the whole of Microsoft Windows Common Controls found 
in mscomctl.ocx. The slider control OCX file is shipped with this package.

* CCRSlider.ocx

This OCX will reside in the program folder. The program reference to this OCX is 
contained within the supplied resource file Panzer Stopwatch Gauge.RES. It is 
compiled into the binary.

* OLEGuids.tlb

This is a type library that defines types, object interfaces, and more specific 
API definitions needed for COM interop / marshalling. It is only used at design 
time (IDE). This is a Krool-modified version of the original .tlb from the 
vbaccelerator website. The .tlb is compiled into the executable.
For the compiled .exe this is NOT a dependency, only during design time.

From the command line, copy the tlb to a central location (system32 or wow64 
folder) and register it.

COPY OLEGUIDS.TLB %SystemRoot%\System32\
REGTLIB %SystemRoot%\System32\OLEGUIDS.TLB

In the VB6 IDE - project - references - browse - select the OLEGuids.tlb



* SETUP.EXE - The program is currently distributed using setup2go, a very useful 
and comprehensive installer program that builds a .exe installer. Youll have to 
find a copy of setup2go on the web as it is now abandonware. Contact me
directly for a copy. The file "install PzStopwatch 0.1.0.s2g" is the configuration 
file for setup2go. When you build it will report any errors in the build.

* HELP.CHM - the program documentation is built using the NVU HTML editor and 
compiled using the Microsoft supplied CHM builder tools (HTMLHelp Workshop) and 
the HTM2CHM tool from Yaroslav Kirillov. Both are abandonware but still do
the job admirably. The HTML files exist alongside the compiled CHM file in the 
HELP folder.

 Project References:

  VisualBasic for Applications  
  VisualBasic Runtime Objects and Procedures  
  VisualBasic Objects and Procedures  
  OLE Automation  
  vbRichClient6  


LICENCE AGREEMENTS:

Copyright © 2023 Dean Beedell

In addition to the GNU General Public Licence please be aware that you may use 
any of my own imagery in your own creations but commercially only with my 
permission. In all other non-commercial cases I require a credit to the 
original artist using my name or one of my pseudonyms and a link to my site. 
With regard to the commercial use of incorporated images, permission and a 
licence would need to be obtained from the original owner and creator, ie. me.

![panzer-clock-ywidget-displa](https://github.com/yereverluvinunclebert/Panzer-StopWatch-Clock-VB6/assets/2788342/daf48cf7-d9fa-4026-84e7-aef518ab72bf)

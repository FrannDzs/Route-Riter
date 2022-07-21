                                MSTS Friction Calculator
                                      Version 2.0
                                 by Joseph T. Realmuto
                                     July 27, 2004

PLEASE READ THIS DOCUMENT IN ITS ENTIRETY, PARTICULARLY THE TERMS AND CONDITIONS NEAR THE END!!

After unzipping the file fcalc_20.zip you should have 23 files:


FILE LIST:

File_id.diz                             Brief description of this package

Fcalc_20.exe                            The friction calculator program

Fcalc_20.c                              Source code for the friction calculator program

Fcalc_20w.exe                           The friction calculator program (Win32 console version)

Fcalc_20w.c                             Source code for the Win32 console friction calculator program

Readme.txt                              This file

FCalc 2.0.gif                           FCalc 2.0 logo

Methodology and Theory of FCalc 2.doc   Microsoft Word document describing how I derived the new
                                        resistance calculation method (requires at least MSWord '97)

Methodology and Theory of FCalc 2.txt   Text only document describing how I derived the new resistance
                                        calculation method (NO CHARTS- only for those without MS Word '97 or better)

FCalc 2.0 Inputs.txt                    Summary of inputs used in FCalc 2.0 source code

FCalc 2.0 Variable Parameters.txt       Summary of variable equations used in FCalc 2.0 source code

Autorack.xls                            Excel spreadsheet for autoracks

COFC.xls                                Excel spreadsheet for container-on-flat cars

Diesel or Electric Locomotive.xls       Excel spreadsheet for diesel and electric locomotives

Empty Hopper.xls                        Excel spreadsheet for empty hoppers

High Speed Train.xls                    Excel spreadsheet for high-speed trains

Passenger Car.xls                       Excel spreadsheet for passenger cars

Railcar(EMU or DMU).xls                 Excel spreadsheet for railcars

Spine Car.xls                           Excel spreadsheet for spine cars

Standard Freight.xls                    Excel spreadsheet for standard freight cars

Steam Locomotive.xls                    Excel spreadsheet for steam locomotives

FCalc 2.0 TOFC.xls                      Excel spreadsheet for trailer-on-flat cars

Muliple Equation Plot.xls               Excel spreadsheet to plot and compare several FCalc values

These files will be collectively referred to elsewhere in this document as "the package".

DISCLAIMER: I am providing this program and the spreadsheets free of charge and with no guarantees whatsoever.  Since I cannot control how they will be used, I bear no responsiblity whatsoever for any real or imagined damages that they may cause to your computer system, or any incidental or consequential damages resulting from any damage to said computer system.  Furthermore, I do not consider myself obliged to provide any technical support other than the documentation included in this package.


USING THE FCALC_20.EXE PROGRAM:

This program is used to calculate friction parameters for new trains so that, hopefully, as people make add-on trains they will also have more prototypical performance.  I also encourage those who make updates to .eng and .wag files for other reasons to also use the more accurate friction parameters calculated by my program.

The program will run under DOS, or in a DOS box under Windows.  The Windows console program works in a similar manner to the DOS version, and should run on systems which can't run DOS programs.  For either version, you will need to know the weight(metric tons), frontal area(square meters), and number of axles of the car or locomotive, as well as the drag coefficient(for locomotives only), and in some cases the length.  The program will prompt you line by line to enter these values after it prompts you for the type of rolling stock, and will then output a line similar to below (program may take several seconds to calculate):

		878N/m/s		-0.10		 1mph		3.47N/m/s		1.745

Replace the first line in the .eng or .wag file under the heading Friction with this new line.   See examples below:

  ORIGINAL LINES:

   Friction (
		100N/m/s		1		-1mph		0		1
		5.1N/rad/s		1		-1rad/s		0		1
	)

  MODIFIED LINES:

	Friction (
		878N/m/s		-0.10		 1mph		3.47N/m/s		1.745
		5.1N/rad/s		1		-1rad/s		0		1
	)

Note that we only replace the first set of parameters.

For more information about how these calculations work and other interesting information about my modifications refer to the file Methodology and Theory of FCalc 2.doc.


TERMS AND LIMITATIONS:

1)You may use the programs Fcalc_20.exe and Fcalc20w.exe freely, and may send copies
  to others so long as this file(Readme.txt) and the source code(Fcalc_20.c and Fcalc20w.c)
  are included.  I especially encourage developers of new and reskinned rolling stock to
  use these programs so that this new rolling stock will have the correct resistance parameters,
  and do not require that you give me any credit when you do so in order to facilitate this.

2)You may use the source code as part of another program provided I am given credit for
  the idea and the use of my source code.  I also require that you send me the final
  version of your program and source code so that I can verify that it does the
  calculations correctly.

3)I bear no responsiblity for any damages, either consequential or inconsequential,
  resulting from the use of anything included in this package.  Use at your own risk!

4)You may NOT post any portion of any part of this package on any Internet sites,
  software, written publications, television or radio programs without my express written
  permission.  IF YOU DO SO, YOU WILL BE SUBJECT TO LITIGATION AND CRIMINAL PENALTIES
  WHICH I WILL VIGOROUSLY PURSUE.  I generally support the free exchange of information
  and ideas on the Internet, but I must require you to ask for permission because
  there are far too many people who post other people's work on their web sites as
  their own in order to increase their page views(and consequently, their income),
  leaving the original author with no credit and possible lost income. 

5)If I have given you permission to use some or all of the materials contained herein on
  Internet sites, software, written publications, television or radio programs you will be
  subject to the following additional conditions:

    a)I must be given credit for my work on the same page that it appears.  However,
      you may NOT post my e-mail address unless I have given you permission to do so.

    b)If the site is paid for by advertisers based on page views, I must be paid a
      previously agreed amount each time the page containing my material is viewed.
      In most cases if the site is non-profit I will allow my work to be posted without
      payment, but never assume anything until you ask me.

    c)If the Fcalc program or documentation is included as part of a future train simulator
      or other type of software, I must be paid an agreed upon amount per unit sold if the
      software is not free.  In cases where the software is free, I may or may not require
      payment, but I still require my permission for FCalc to be included.

    d)You may only use the specific materials for which you have requested and been
      granted permission to use, and nothing else.

    e)I may withdraw my permission if any of the conditions in 5a, 5b, 5c, and 5d are
      violated, OR if the material is used on a site which contains pornographic or
      other obscene material, or has links to sites containing pornographic or other
      obscene material.  If I withdraw my permission, you will be given 30 days notice
      to either comply with the original terms of the agreement, or remove my material
      from the site.  If you fail to do either, I will pursue civil and criminal
      penalties against you to the fullest extent of the law.


CONTACT INFORMATION:

PM @ train-sim.com
PM @ 3DTrains.com

e-mail: jtr1962@yahoo.com 

Please use the e-mail address to contact me if you want permission to put anything in this package on an Internet site, or if you are a programmer and wish to send me a finished copy of your work which incorporates my source code, or for anything else that requires the exchange of files.  Use the personal message system of either train-sim.com or 3DTrains.com for any other questions.


CREDITS (alphabetical order):

Kevin Arceneaux-beta testing
Frank Bandre-beta testing
Bob Boudoin-empirical data for many types of freight equipment, extensive physics modifications, extensive beta testing
Richard Gibb-beta testing
Todd Jones-locomotive engineer, empirical data for many types of freight equipment
Derek Morton-programmer of EngMod, volunteered to program future GUI version of FCalc 2
Dave Nelson-solid bearing equations, extensive help incorporating solid bearings into Fcalc, help with spreadsheets
Cyndi Richards-beta testing
Clem Tillier-data on high-speed trains
James Titus-ideas to incorporate solid bearings into Fcalc, help with spreadsheets
Richard Wilms-anecdotal information about solid (friction) bearings

I've included anyone to whom I sent a pre-release copy of Fcalc 2 on this list as a beta tester.  I've also included anyone else whom I'm aware was involved in beta testing.  If I've missed anyone, please let me know.

Finally, I want to express thanks to the entire community for making Fcalc 1 such an overwhelming success and accepting it as the defacto standard!
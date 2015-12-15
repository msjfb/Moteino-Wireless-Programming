# Moteino-Wireless-Programming

Version 1.4.1 (2015/12/15)
Corrected problem in error trapping routine: did not display the correct information.
===============================================================
Version 1.4 (2015/12/09)
Added support for use of a TCP port to communicate with the Gateway. 
This allows the gateway to be connected anywhere on the network using a Device Server (such as the Lantronix UDS-1000) 
or a PC with a "Serial to Ip" software running (such as the free Serial-TCP program)

===============================================================
Version 1.2 (2015/11/14)
Added command line support for unattended mode: 4 parameters are required:
<COM port>,<BaudRate>,<TargetNode>,<Filepath>

Ex: MoteinoWirelessProgramming 6,115200,5,c:\Arduino\Hexfiles\MyHexfile.HEX
This would launch the program in minimized mode, using COM port #6 at 115200 baud, and transmitting file "c:\Arduino\Hexfiles\MyHexfile.HEX"

The program will automatically exit when finished with the following return codes that can be read from a .BAT or .CMD file for automatic processing.
Code   Description
0   Success
20   Cannot open COM port
30   Cannot set Target on Gateway
40   Handshake NAK (Image refused by target)
50   Problem in processing HEX file
60   HEX file not found

===============================================================
Version 1.1  (2015/11/02)
Saves/reads last HEX file used in MoteinoWP.ini file

===============================================================
Version 1.0
Initial release

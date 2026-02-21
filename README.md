# Lighthouse

Lighthouse is a Python program which provides a GUI for collecting simultaneous electrochemical and spectroscopic data using a [Gamry](https://www.gamry.com/) potentiostat and [OceanOptics ](https://www.oceanoptics.com/)spectrometer. Gamry hardware is controlled with their proprietary [Python library](https://www.gamry.com/support-2/software/echem-toolkitpy/), which you can only obtain directly from them, while the OceanOptics hardware is controlled with [Seabreeze](https://github.com/ap--/python-seabreeze). It is possible that future releases will support other hardware as time allows.

Lighthouse was originally created by Carter Pryor for the [Graham group](https://graham.as.uky.edu/) at the University of Kentucky.

### Installation

Running Lighthouse requires a Python installation with numpy, matplotlib, Seabreeze, and Gamry's proprietary Echem toolkitpy library. At the moment, the only way to obtain a copy of toolkitpy is with Gamry's installer for its software packages (the same installer used for Framework, EChem Analyst, etc.). Because of this, it is necessary to install the Gamry software first, and then install the other dependencies. 

In more detail:

1. 

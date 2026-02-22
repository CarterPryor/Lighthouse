# Lighthouse

Lighthouse is a Python program which provides a GUI for collecting simultaneous electrochemical and spectroscopic data using a [Gamry](https://www.gamry.com/) potentiostat and [OceanOptics ](https://www.oceanoptics.com/)spectrometer. Gamry hardware is controlled with their proprietary [Python library](https://www.gamry.com/support-2/software/echem-toolkitpy/), which you can only obtain directly from them, while the OceanOptics hardware is controlled with [Seabreeze](https://github.com/ap--/python-seabreeze). It is possible that future releases will support other hardware as time allows.

At the moment, Lighthouse is only configured to run cyclic voltammetry (CV) electrochemistry experiments, but there are plans to eventually allow for a variety of other electrochemical measurements - chronopotentiometry, chronoamperometry, open circuit voltage, etc.

Lighthouse was originally created by Carter Pryor for the [Graham group](https://graham.as.uky.edu/) at the University of Kentucky.

### Installation

Running Lighthouse requires a Python installation with numpy, matplotlib, Seabreeze, and Gamry's proprietary Echem toolkitpy library. At the moment, the only way to obtain a copy of toolkitpy is with Gamry's installer for its software packages (the same installer used for Framework, EChem Analyst, etc.). Because of this, it is necessary to install the Gamry software first, and then install the other dependencies. 

In more detail:

1. Download the installer for the Gamry software package. If your Gamry instrument is registered to your account on Gamry's website, you should be able to download the latest version after signing in. If you had issues with registering your instrument, you may need to contact Gamry support and ask for the latest installer.

2. Run the installer, being sure to tick the "Toolkitpy" box when asked which parts of the Gamry software package you want to install.

3. If you ran a recent version of the installer (I tested with 7.11.0.17837), this should complete creating a Python installation with Gamry's library. It should be located in the folder you chose to install Gamry in, under `Python/PythonXX-XX` (the Xs are digits in a Python version number)

4. Now we need to install the other dependencies into this Python installation. Open a terminal window in the `Python/PythonXX-XX` folder (in Windows 11: Right click -> Open in Terminal), then run the command: 
   
   `.\python.exe -m pip install numpy matplotlib seabreeze`
   
   Once this finishes, your Python installation is ready to run Lighthouse.

5. Clone the repository onto your machine. If you're looking at the Github page for this project, in the top right, click Code -> Download ZIP. In later versions, we may have official releases that are more stable than the current version of the repo.

6. Once you've extracted the folder, you can run Lighthouse as you would any Python program, only being careful to use the interpreter from the Gamry installation since this has the necessary libraries. For convenience's sake, it may be helpful to [create a virtual environment](https://docs.python.org/3/tutorial/venv.html) that references the Gamry installation, especially if you're running multiple Python versions on your machine or if you are running the program from within an IDE like VS Code. 

7. (Optional) Finally, if you wish to use the email features of Lighthouse, you need to create a file called `auth_token.txt` in Lighthouse's directory. The first line of this file should be a valid GMail account, and the second line should be an auth token you generate for that account. You can use other email addresses too, you just need to modify the two lines of code in `./SpecEChemProgram.py` which use the Google SMTP server. 

### Usage

Lighthouse is designed to be intuitive and easy to use without instructions, but if you need additional guidance, please check our [Wiki](https://github.com/CarterPryor/Lighthouse/wiki). 

### Bugs and Issue Reporting

Lighthouse is provided as-is, without warranty. We want to encorporate as many features and safeguards to proper data collection and safety, but as with anything in science and research, we highly encourage you to take your own additional measures to ensure your experiments can proceed uninterrupted and that your data is safe. 

All that being said, if you encounter an issue while using Lighthouse, please let us know! You can report an issue on the [issue tracker on Github](https://github.com/CarterPryor/Lighthouse/issues). This makes sure that both the developer and other users are aware of the problem.  

# Graphical program for collecting electrochemical and spectral data simultaneously using a Gamry potentiostat and OceanOptics spectrometer
# Originally written by Carter Pryor (carter_pryor@outlook.com) for Graham group at UKY
# Last modified 2026-06-12

# Recommended sample period: >=0.1 s according to Gamry docs

# Run command:
# & 'C:\Program Files (x86)\Gamry Instruments\Python\Python37-32\python.exe' .\SpecEChemProgram.py

# Lightbulb icon source: <a href="https://www.flaticon.com/free-icons/idea" title="idea icons">Idea icons created by Good Ware - Flaticon</a>

'''
TODO:
 - Update email sending to include whether or not expt finished completely
- Warning if user pstat settings are over max number of data points
- fix the bounds of each plotting window, and only change it when user asks
- Change the threading model to avoid worker threads pushing GUI updates (plotting)
- Break up code into multiple files to increase readability

- Fix bug preventing running too long experiments (sequence wizard type thing)

- Test everything flushes every 30 s correctly

- Consider additionally snapping the data collection to the start of each cycle to improve stability
- Allow user to choose the maximum current before automatic stop
- Make an option (either another program or an option in this one) for chronocoulometry and perhaps for open circuit measurements
'''

# Import Python's built-in functionality for things like timing, etc.
import datetime
import time
# And for multithreading
import threading
# and for logging
import logging
import pathlib
import math

# Import Python's built-in GUI library
import tkinter as tk
import tkinter.messagebox as mbox
import tkinter.ttk
import tkinter.simpledialog as simpledialog
import tkinter.filedialog as filedialog

# Import numpy, for math
import numpy as np
# and matplotlib, for plotting
import matplotlib.figure as figure
from matplotlib.backends.backend_tkagg import FigureCanvasTkAgg

# libraries for sending notification emails when complete
import smtplib
import ssl
from email import encoders
from email.mime.base import MIMEBase
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
# create zip file to send data
import zipfile

# Import Seabreeze, the library which controls an OceanOptics spectrometer
import seabreeze as sb
import seabreeze.spectrometers

# Gamry's toolkitpy library for controlling the Gamry pstat directly from Python
import toolkitpy as tkp

# initialize logger
logger = logging.getLogger(__name__)

def perf_sleep_until(final_time: int):
    while True:
        now = time.perf_counter_ns()
        dt = final_time - now
        # if we are at time or past it, stop sleeping
        if (dt <= 0):           
            break
        # if we are more than 200 ms away from time, sleep for 80% of remaining time
        elif (dt > 0.200*(1E9)):
            time.sleep(dt * 0.8 * (1E-9))
        # if we are more than 50 ms away from time, sleep for 15 ms
        elif (dt > 0.050*(1E9)):
            time.sleep(0.015)
        # otherwise, just continue the loop
        else:
            # yield to other processes; this can reduce CPU jitter and make this more accurate
            time.sleep(0)

# A window class, which holds all the information for the our active experiment
class MyWindow:

    # Constructor, called when we make a new object of this class
    # Sets up the GUI and the functions that will be called when we click on parts of it
    def __init__(self):
        # Initialize Toolkitpy for this file
        tkp.toolkitpy_init("Lighthouse_SpecEChemProgram.py")

        # Set up the window
        self.root = tk.Tk()
        self.root.geometry("1000x600")
        self.root.title("Lighthouse")

        # attach the function that runs before exiting, confirming we want to quit
        self.root.protocol("WM_DELETE_WINDOW", self.confirm_quit)
        # for now, make window not resizeable until I change the plot updates to be in main thread
        self.root.resizable(False, False)

        # Then, add all the pieces to the GUI

        ### Top level GUI code

        # Menu bar
        self.menubar = tk.Menu(self.root)
        # File submenu
        self.menu_file = tk.Menu(self.menubar, tearoff=0)
        self.menu_file.add_command(label="Choose output folder", command=self.choose_out_dir)
        # Edit submenu
        self.menu_edit = tk.Menu(self.menubar, tearoff=0)
        self.menu_edit.add_command(label="Experiment Title", command=self.edit_exp_name)
        self.menu_edit.add_command(label="Operator", command=self.edit_operator)
        self.menu_edit.add_command(label="Description", command=self.edit_description)
        self.menu_edit.add_command(label="Notification emails", command=self.edit_emails)
        # Add the menus to the menubar
        self.menubar.add_cascade(label="File", menu = self.menu_file)
        self.menubar.add_cascade(label="Edit", menu = self.menu_edit)
        # Set the menu of the window to be the menu we made
        self.root.config(menu = self.menubar)

        # Filename label
        self.lbl_filename = tk.Label(self.root, text="Save to: default")
        self.lbl_filename.place(x=20, y=10)
        # actively running label
        self.lbl_running = tk.Label(self.root, text="not running", fg="red")
        self.lbl_running.place(x=850, y=560)
        # Checkbox to include or not include potentiostat in the current measurement
        self.use_pstat = tk.BooleanVar()
        self.chk_use_pstat = tk.Checkbutton(self.root, text="Use Potentiostat", variable=self.use_pstat, onvalue=True, offvalue=False)
        self.chk_use_pstat.select()
        self.chk_use_pstat.place(x=20, y=540)
        # Checkbox to include or not include spectrometer in current measurement
        self.use_spec = tk.BooleanVar()
        self.chk_use_spec = tk.Checkbutton(self.root, text="Use Spectrometer", variable=self.use_spec, onvalue=True, offvalue=False)
        self.chk_use_spec.select()
        self.chk_use_spec.place(x=20+20+470, y=540)
        # Button to actually start the measurement
        self.btn_start = tk.Button(self.root, text="Start!", command=self.start_measurement)
        self.btn_start.place(x=750, y=560)
        # Button to abort measurement
        self.btn_abort = tk.Button(self.root, text="Abort", command=self.abort_measurement)
        self.btn_abort.place(x=800, y=560)

        # Default experiment name & operator
        self.experiment_name = "Experiment"
        self.operator = "Graham Lab"
        self.description = ""
        self.emails = ""

        # Default save directory = current directory:
        self.save_dir = "."

        # Bool for if the experiment is actively running
        self.running = False

        ### End top level GUI code

        ### PStat related init code

        # The frame for potentiostat-related settings
        self.frame_pstat = tk.Frame(self.root, width=470, height=500, bd=2, relief="sunken")
        self.frame_pstat.place(x=20, y=40)

        # if pstat is connected label
        self.lbl_pstat = tk.Label(self.frame_pstat, text="Potentiostat:")
        self.lbl_pstat.place(x=1, y=1)
        self.lbl_pstat_connected = tk.Label(self.frame_pstat, text="Not connected", fg="red")
        self.lbl_pstat_connected.place(x=70, y=1)
        # pstat model label
        self.lbl_pstat_model = tk.Label(self.frame_pstat, text="Model: ---")
        self.lbl_pstat_model.place(x=1, y=20)
        # cell off/on label
        self.lbl_cell_state = tk.Label(self.frame_pstat, text="Cell Off", fg="red")
        self.lbl_cell_state.place(x=1, y=40)
        # button to connect PStat
        self.btn_connect_pstat = tk.Button(self.frame_pstat, text="Connect Potentiostat", command=self.connect_pstat)
        self.btn_connect_pstat.place(x=250, y=20)
        # pstat potential reading label
        self.lbl_pstat_potential = tk.Label(self.frame_pstat, text="E: --- V", font=("Arial", 11))
        self.lbl_pstat_potential.place(x=150, y=80)
        # pstat current reading label
        self.lbl_pstat_current = tk.Label(self.frame_pstat, text="i: --- A", font=("Arial", 11))
        self.lbl_pstat_current.place(x=250, y=80)

        # All this GUI stuff is for choosing settings for a given experiment

        # First vertex potential: where we start the scan
        # Label
        self.lbl_vertex_pot1 = tk.Label(self.frame_pstat, text="Vertex Potential 1 (V):")
        self.lbl_vertex_pot1.place(x=20, y=120)
        # Text box
        self.vertex_pot1_text = tk.StringVar()
        self.entry_vertex_pot1 = tk.Entry(self.frame_pstat, textvariable=self.vertex_pot1_text)
        self.entry_vertex_pot1.place(x=145, y=122, width=35)

        # Vertex potential 2: where we turn around in the cycles
        # Label
        self.lbl_vertex_pot2 = tk.Label(self.frame_pstat, text="Vertex Potential 2 (V):")
        self.lbl_vertex_pot2.place(x=20, y=140)
        # Text box
        self.vertex_pot2_text = tk.StringVar()
        self.entry_vertex_pot2 = tk.Entry(self.frame_pstat, textvariable=self.vertex_pot2_text)
        self.entry_vertex_pot2.place(x=145, y=142, width=35)

        # Scan rate
        # Label
        self.lbl_scan_rate = tk.Label(self.frame_pstat, text="Scan Rate (V/s):")
        self.lbl_scan_rate.place(x=20, y=160)
        # Text box
        self.scan_rate_text = tk.StringVar()
        self.entry_scan_rate = tk.Entry(self.frame_pstat, textvariable=self.scan_rate_text)
        self.entry_scan_rate.place(x=145, y=162, width=35)

        # Step size
        # Label
        self.lbl_step_size = tk.Label(self.frame_pstat, text="Sample Time (s):")
        self.lbl_step_size.place(x=200, y=120)
        # Text box
        self.step_size_text = tk.StringVar()
        self.entry_step_size = tk.Entry(self.frame_pstat, textvariable=self.step_size_text)
        self.entry_step_size.place(x=295, y=122, width=35)

        # Number of cycles
        # Label
        self.lbl_num_cycles = tk.Label(self.frame_pstat, text="# Cycles:")
        self.lbl_num_cycles.place(x=200, y=140)
        # Text box
        self.num_cycles_text = tk.StringVar()
        self.entry_num_cycles = tk.Entry(self.frame_pstat, textvariable=self.num_cycles_text)
        self.entry_num_cycles.place(x=295, y=142, width=35)

        self.figure_cv = figure.Figure((4.55, 3), dpi=100)
        self.axes_cv = self.figure_cv.subplots()
        self.canv_cv = FigureCanvasTkAgg(self.figure_cv, self.frame_pstat)
        self.canv_cv.get_tk_widget().place(x=6, y=190)

        # Variables that keep the program from doing things it shouldn't when it shouldn't
        self.has_potentiostat = False
        self.should_draw_pstat = False
        # Potentiostat handle
        self.potentiostat = None

        ### End PStat related init code

        ### Spectrometer-related init code

        # The frame for spectrometer-related settings
        self.frame_spec = tk.Frame(self.root, width=470, height=500, bd=2, relief="sunken")
        self.frame_spec.place(x=20+20+470, y=40) 
        
        # if spectrometer is connected label
        self.lbl_spectrometer = tk.Label(self.frame_spec, text="Spectrometer:")
        self.lbl_spectrometer.place(x=1, y=1)
        self.lbl_spec_connected = tk.Label(self.frame_spec, text="Not connected", fg="red")
        self.lbl_spec_connected.place(x=80, y=1)
        # spectrometer model label
        self.lbl_spec_model = tk.Label(self.frame_spec, text="Model: ---")
        self.lbl_spec_model.place(x=1, y=20)
        # Connect button
        self.btn_connect_spec = tk.Button(self.frame_spec, text="Connect Spectrometer", command=self.connect_spectrometer)
        self.btn_connect_spec.place(x=270, y=10)
        # integration time label
        self.lbl_integ_time = tk.Label(self.frame_spec, text="Integration Time:", font=("Arial", 11))
        self.lbl_integ_time.place(x=1, y=50)
        # textbox for integration time entry
        self.integ_time_txt = tk.StringVar()
        self.entry_integ_time = tk.Entry(self.frame_spec, textvariable=self.integ_time_txt)
        self.entry_integ_time.place(x=121, y=53, width=35)
        # ms label
        self.lbl_ms = tk.Label(self.frame_spec, text="ms")
        self.lbl_ms.place(x=156, y=50)
        # Set integration time from entry label
        self.btn_integ_time = tk.Button(self.frame_spec, text="Set", command=self.set_integ_time)
        self.btn_integ_time.place(x=180, y=50)
        # electric dark correction checkbox
        self.enable_dark_correction = tk.BooleanVar()
        self.chk_dark = tk.Checkbutton(self.frame_spec, text="Electric Dark Correction", variable=self.enable_dark_correction, onvalue=True, offvalue=False)
        self.chk_dark.select() # check the box
        self.chk_dark.place(x=1, y=80)
        # nonlinearity correction checkbox
        self.enable_nonlinearity_correction = tk.BooleanVar()
        self.chk_nonlin = tk.Checkbutton(self.frame_spec, text="Nonlinearity Correction", variable=self.enable_nonlinearity_correction, onvalue=True, offvalue=False)
        self.chk_nonlin.select()
        self.chk_nonlin.place(x=150, y=80)
        # Store reference spectrum button
        self.ref_spec_photo = tk.PhotoImage(file="lightbulb-on-ico.png")
        self.btn_store_ref_spec = tk.Button(self.frame_spec, image=self.ref_spec_photo, command=self.store_reference_spectrum)
        self.btn_store_ref_spec.place(x=139, y=115)
        # Store dark spectrum button
        self.dark_spec_photo = tk.PhotoImage(file="lightbulb-off-ico.png")
        self.btn_store_dark_spec = tk.Button(self.frame_spec, image=self.dark_spec_photo, command=self.store_dark_spectrum)
        self.btn_store_dark_spec.place(x=299, y=115)
        # Canvas to draw currently measured spectrum on
        self.figure_spectrum = figure.Figure((4.55, 3), dpi=100)
        self.axes_spectrum = self.figure_spectrum.add_subplot()
        self.canv_spectrum = FigureCanvasTkAgg(self.figure_spectrum, self.frame_spec)
        # right-click context menu for the spectrum
        self.menu_canv_spectrum = tk.Menu(self.canv_spectrum.get_tk_widget(), tearoff=0)
        self.menu_canv_spectrum.add_command(label="Set y-min", command=self.canv_spectrum_set_ymin)
        self.menu_canv_spectrum.add_command(label="Set y-max", command=self.canv_spectrum_set_ymax)
        self.canv_spectrum.get_tk_widget().bind("<Button-3>", self.canv_spectrum_popup)
        self.canv_spectrum.get_tk_widget().place(x=6, y=160)

        # variable that stores if we should be repeatedly looping drawing the spectrum
        self.should_draw_spec = False
        self.should_reset_spec_limits = False # flag that indicates that spectrometer plot limits need to be reset
        # variables storing the limits for each plotting mode besides 
        self.spec_plot_ylims_sub = (0, 180000)
        self.spec_plot_ylims_t = (-0.1, 1.1)
        self.spec_plot_ylims_abs = (-0.1, 2)

        # Label and box to input how often we collect spectra
        self.lbl_collect = tk.Label(self.frame_spec, text="Collect every:")
        self.lbl_collect.place(x=70, y=470)
        self.spec_freq_txt = tk.StringVar()
        self.entry_collect_frequency = tk.Entry(self.frame_spec, textvariable=self.spec_freq_txt)
        self.entry_collect_frequency.place(x=147, y=472, width=30)
        # Unit selection list box
        self.spec_freq_units = tk.StringVar() # Variable to capture input from the frequency combo box
        self.combo_spec_freq_units = tk.ttk.Combobox(self.frame_spec, width=4, textvariable=self.spec_freq_units)
        self.combo_spec_freq_units["values"] = ("ms", "s", "min", "hr") # Possible options to select from
        self.combo_spec_freq_units.current(0) # make ms the default unit
        self.combo_spec_freq_units.state(["readonly"]) # User can only pick from pre-defined options
        self.combo_spec_freq_units.place(x=183, y=471)
        # How should intensities be saved?
        # Each file will always include the saved reference and dark spectra, so you can always go back and forth, but
        # this setting controls what it defaults to
        self.lbl_as = tk.Label(self.frame_spec, text="as")
        self.lbl_as.place(x=230, y=470)
        self.spec_intensity_type = tk.StringVar()
        self.combo_spec_intensity_type = tk.ttk.Combobox(self.frame_spec, width=10, textvariable=self.spec_intensity_type)
        self.combo_spec_intensity_type["values"] = ("Raw Int.", "Raw Int. - Ref", "%T or %R", "Abs")
        # Callback to trigger redraws if the box is selected
        self.combo_spec_intensity_type.bind("<<ComboboxSelected>>", self.combo_spec_intensity_changed)
        # Raw intensity = raw spectrometer data
        # Raw Int. - Ref = Raw spectrometer data minus the saved reference spectrum
        # %T or %R = (Raw spectrum - dark spectrum) / (Ref. Spectrum - dark spectrum)
        # Abs = -log ([Raw spectrum - dark spectrum]/ [Ref. spectrum - dark spectrum])
        self.combo_spec_intensity_type.current(0)
        self.combo_spec_intensity_type.state(["readonly"])
        self.combo_spec_intensity_type.place(x=248, y=471)
        # Collect now button
        self.btn_collect_spec = tk.Button(self.frame_spec, text="Collect now", command=self.collect_spec_now)
        self.btn_collect_spec.place(x=350, y=470)

        self.has_spectrometer = False
        self.spectrometer = None
        self.reference_spec = None
        self.dark_spec = None
        # a lock on accessing the spectrometer between threads
        self.spectrometer_lock = threading.Lock()

        ### End spectrometer related initialization code

    # function to confirm if the user wants to quit while an experiment is running
    def confirm_quit(self):
        # if experiment is running
        if self.running == True:
            # ask if the user wants to quit
            response = mbox.askyesno("Confirm Quit", "An experiment is currently running. Are you sure you want to close the software?")
            # If they say yes, close
            if response == True:
                # Abort measurement
                self.abort_measurement()
                # Wait for measurement thread to finish clean-up
                self.thread_measurement.join()
                # end all the other threads if necessary
                if (hasattr(self, "thread_draw_pstat")):
                    self.should_draw_pstat = False
                    self.thread_draw_pstat.join()
                # then close window
                self.root.destroy()
            # otherwise, do nothing
        # if experiment is not running, just close
        else:
            # stop running all threads then let them join
            if (hasattr(self, "thread_draw_pstat")):
                self.should_draw_pstat = False
                self.thread_draw_pstat.join()
            self.root.destroy()

    # update all labels
    def gui_update(self):
        logger.debug("GUI Update loop top")
        # expt running & cycle label
        if (self.running):
            # if we don't have a first pstat point yet, we're still at cycle 0
            if (hasattr(self, "most_recent_pstat_pt")):
                now_cycle = self.most_recent_pstat_pt[8]
            else:
                now_cycle = 0
            # using current_expt_max_cycles shouldn't be a race condition as long as you only set
            # running = True AFTER setting current_expt_max_cycles for each new expt run
            logger.debug(f"Updating running label to cycle: {now_cycle}")
            self.lbl_running.configure(text=f"Running: Cycle {now_cycle}/{self.current_expt_max_cycles}", fg="green")
        else:
            logger.debug("Updating running label to not running")
            self.lbl_running.configure(text="not running", fg="red")

        # Updates that only happen if spectrometer is connected (empty right now)

        # Updates that only happen if pstat is connected
        if (self.has_potentiostat == True):
            # cell on/off label
            logger.debug("Updating cell state label")
            if (self.potentiostat.cell() == tkp.CELLSTATE.CELL_ON):
                self.lbl_cell_state.configure(text="Cell On", fg="green")
            else:
                self.lbl_cell_state.configure(text="Cell Off", fg="red")
            # voltage/current labels
            if (self.running): 
                logger.debug("Getting E/i from experiment most recent pt")
                # calling "measure" directly may autorange the pstat which we don't want during a measurement
                # so while we are running, we just use the most recent pstat data point to update labels
                potential = self.most_recent_pstat_pt[2]
                current = self.most_recent_pstat_pt[4]
            else:
                logger.debug("Getting E/i from measure_x() func")
                potential = self.potentiostat.measure_v()
                current = self.potentiostat.measure_i()
            
            logger.debug("Updating pstat potential & current labels")
            # update labels
            self.lbl_pstat_potential.configure(text=f"E: {potential:.3f}V")
            self.lbl_pstat_current.configure(text=f"i: {current:.3E} A") # the .3E = 3 decimal places, in scientific notation
            logger.debug("GUI Update bottom")

        # loop GUI update
        logger.debug("GUI Update loop bottom")
        self.root.after(500, self.gui_update)

    # Open a dialogue box to change the experiment name
    def edit_exp_name(self):
        self.experiment_name = simpledialog.askstring("Edit Experiment Title", "Enter the experiment title (letters, numbers, spaces, and _ OK): ", initialvalue=self.experiment_name)

    # Open dialogue box to change the operator name
    def edit_operator(self):
        self.operator = simpledialog.askstring("Edit Operator Name", "Enter Operator Name: ", initialvalue=self.operator)
    
    def edit_description(self):
        self.description = simpledialog.askstring("Edit Description", "Enter experiment description: ", initialvalue=self.description)

    def edit_emails(self):
        self.emails = simpledialog.askstring("Edit Emails", "Enter email(s) to notify when experiment finishes, separated by comma", initialvalue=self.emails)

    # Choose the directory where the files will be outputted (filename is generated based on experiment name)
    def choose_out_dir(self):
        self.save_dir = filedialog.askdirectory(initialdir=self.save_dir)
        self.lbl_filename.configure(text=f"Save to: {self.save_dir}")

    def combo_spec_intensity_changed(self, event):
        # trigger a reset of the spec limits
        if (self.has_spectrometer == True):
            self.should_reset_spec_limits = True

    def canv_spectrum_popup(self, event):
        try:
            self.menu_canv_spectrum.tk_popup(event.x_root, event.y_root)
        finally:
            self.menu_canv_spectrum.grab_release()
    
    def canv_spectrum_set_ymin(self):
        # do nothing if we aren't plotting
        if (self.has_spectrometer == False):
            return
        
        # ask new ymin
        intensity_type = self.spec_intensity_type.get()
        if (intensity_type == "Raw Int."):
            # we always use the same limits for this type, so do nothing
            return
        elif (intensity_type == "Raw Int. - Ref"):
            old_min, old_max = self.spec_plot_ylims_sub
            new_ymin = simpledialog.askfloat("Spectrometer Plot y-min", "Enter the new y-min", initialvalue=old_min)
            if (new_ymin is not None):
                self.spec_plot_ylims_sub = (new_ymin, old_max)
        elif (intensity_type == "%T or %R"):
            old_min, old_max = self.spec_plot_ylims_t
            new_ymin = simpledialog.askfloat("Spectrometer Plot y-min", "Enter the new y-min", initialvalue=old_min)
            if (new_ymin is not None):
                self.spec_plot_ylims_t = (new_ymin, old_max)
        elif (intensity_type == "Abs"):
            old_min, old_max = self.spec_plot_ylims_abs
            new_ymin = simpledialog.askfloat("Spectrometer Plot y-min", "Enter the new y-min", initialvalue=old_min)
            if (new_ymin is not None):
                self.spec_plot_ylims_abs = (new_ymin, old_max)
        # flag spec plot limits for an update
        self.should_reset_spec_limits = True
        

    def canv_spectrum_set_ymax(self):
        # do nothing if we aren't plotting
        if (self.has_spectrometer == False):
            return
        
        # ask new ymin
        intensity_type = self.spec_intensity_type.get()
        if (intensity_type == "Raw Int."):
            # we always use the same limits for this type, so do nothing
            return
        elif (intensity_type == "Raw Int. - Ref"):
            old_min, old_max = self.spec_plot_ylims_sub
            new_ymax = simpledialog.askfloat("Spectrometer Plot y-max", "Enter the new y-max", initialvalue=old_max)
            if (new_ymax is not None):
                self.spec_plot_ylims_sub = (old_min, new_ymax)
        elif (intensity_type == "%T or %R"):
            old_min, old_max = self.spec_plot_ylims_t
            new_ymax = simpledialog.askfloat("Spectrometer Plot y-max", "Enter the new y-max", initialvalue=old_max)
            if (new_ymax is not None):
                self.spec_plot_ylims_t = (old_min, new_ymax)
        elif (intensity_type == "Abs"):
            old_min, old_max = self.spec_plot_ylims_abs
            new_ymax = simpledialog.askfloat("Spectrometer Plot y-max", "Enter the new y-max", initialvalue=old_max)
            if (new_ymax is not None):
                self.spec_plot_ylims_abs = (old_min, new_ymax)
        # flag spec plot limits for an update
        self.should_reset_spec_limits = True

    # Attempt to access the Gamry Potentiostat
    def connect_pstat(self):
        # Get list of connected Gamry devices
        logger.info("Attempting to connect to pstat...")
        device_list = tkp.enum_sections()
        if (len(device_list) == 0): # List is empty, no attatched pstat
            logger.warning("No Gamry Potentiostat found")
            mbox.showwarning("Warning", "No connected Gamry Potentiostats found!")
        else:
            # We have a pstat, try to connect to first one
            logger.info("At least on PStat available, trying to connect...")
            self.potentiostat = tkp.Pstat(device_list[0])
            # Open connection to the Pstat
            self.potentiostat.open() 
            model = self.potentiostat.model_no()
            print(f"Serial No.: {self.potentiostat.serial_no()}")
            logger.info(f"Connected pstat: Serial No. {self.potentiostat.serial_no()}")
            self.has_potentiostat = True
            # TODO: See about changing current convention (which direction is positive?)
            self.lbl_pstat_connected.configure(text="connected", fg="green")
            self.lbl_pstat_model.configure(text=f"Model: {model}")
            # Begin updating the labels showing voltage and current
            self.should_update_pstat_labels = True
            # Legacy code - TBR
            # self.thread_pstat_labels = threading.Thread(target=self.update_pstat_readinglabels)
            # self.thread_pstat_labels.start()
    
    ''' Legacy Code - TBR
    def update_pstat_readinglabels(self):
        # Loop
        while self.should_update_pstat_labels:
            # only run if potentiostat is connected
            if (self.has_potentiostat):
                # if an experiment is actively running, just get the most recent value from that.
                # The measure commands can force the potentiostat to autorange the current, which we
                # don't want during an experiment more than necessary
                if (self.running):
                    potential = self.most_recent_pstat_pt[2]
                    current = self.most_recent_pstat_pt[4]
                else:
                    potential = self.potentiostat.measure_v()
                    current = self.potentiostat.measure_i()

                self.lbl_pstat_potential.configure(text=f"E: {potential:.3f}V")
                self.lbl_pstat_current.configure(text=f"i: {current:.3e} A")
            else:
                break # stop execution of this thread if pstat no longer connected

            time.sleep(0.5) # sleep for half a second
    '''
    # Attempt to connect to an OceanOptics spectrometer via USB
    def connect_spectrometer(self):
        logger.info("Attempting to connect to OceanOptics Spectrometer...")
        try:
            logger.info("At least one spectrometer available, trying to connect...")
            # Try to grab the first spectrometer available
            self.spectrometer = sb.spectrometers.Spectrometer.from_first_available()
            logger.info("Spectrometer connected successfully")
            # If we get one, update the text to show we are connected and what the spectrometer model label is
            self.lbl_spec_connected.configure(text="Connected", fg="green")
            model = self.spectrometer.model
            logger.info(f"Spectrometer model: {model}")
            self.lbl_spec_model.configure(text=f"Model: {model}")
            
            self.has_spectrometer = True

            # Also update the integration time to be the minimum the instrument supports
            limits = self.spectrometer.integration_time_micros_limits
            # Update the instrument
            self.spectrometer.integration_time_micros(limits[0])
            self.integration_time_micros = limits[0]
            # Update the text box
            min_ms = limits[0] / 1000 # Convert from micros to ms
            self.entry_integ_time.delete(0, tk.END)
            self.entry_integ_time.insert(0, f"{min_ms}")

            # The wavelengths attached to each pixel don't change. So save them at the start to save some processing time
            self.wavelengths = self.spectrometer.wavelengths()

            self.has_reference_spec = False
            self.dark_spec = np.zeros((len(self.wavelengths)))

            # Finally, start drawing the current spectrum the instrument records on the canvas on repeat
            self.draw_first_spec()
        except sb.spectrometers.SeaBreezeError as e:
            logger.warning("No available OceanOptics Spectrometer found")
            logger.warning(e)
            # If we can't get a spectrometer
            self.spectrometer = None
            mbox.showwarning("Failed to connect to OceanOptics Spectrometer", e)
    
    # draw the first spectrum, that will later be modified by repeat calls to draw_spec
    def draw_first_spec(self):
        logger.info("Starting drawing first spectrum")
        now_intensities = self.spectrometer.intensities(self.enable_dark_correction, self.enable_nonlinearity_correction)
        
        # first spectrum is always just raw intensity
        self.axes_spectrum.clear()
        # disable autoscale
        self.axes_spectrum.autoscale(enable=False)
        # plot the raw intensities
        logger.debug("draw_first_spec: calling plot()")
        self.line2d_spec, = self.axes_spectrum.plot(self.wavelengths, now_intensities, color="blue")
        # plot a "max intensity" line at 170,000 intensity
        self.line2d_spec_int_max, = self.axes_spectrum.plot([0, 3000], [170000, 170000], "--", color="red")
        # set plot labels, etc.
        self.axes_spectrum.set_xlabel("Wavelength (nm)")
        self.axes_spectrum.set_ylabel("Intensity (a.u.)")
        self.axes_spectrum.set_xlim(self.wavelengths[0], self.wavelengths[-1])
        self.axes_spectrum.set_ylim(0, 180000)
        self.axes_spectrum.grid()
        # finally execute the draw call
        logger.debug("draw_first_spec: calling draw()")
        self.canv_spectrum.draw()
        # begin the looping draw
        self.root.after(1000, self.draw_spec)

    # Draw command, to draw the most recent spectrum
    def draw_spec(self):
        logger.debug("draw_spec: Beginning of call")
        if (self.has_spectrometer):
            # block access to spectrometer til intensities call returns
            with self.spectrometer_lock:
                now_intensities = self.spectrometer.intensities(self.enable_dark_correction, self.enable_nonlinearity_correction)
            intensity_type = self.spec_intensity_type.get()

            if (intensity_type == "Raw Int."):
                logger.debug("Updating y-data for raw intensity")
                # update line data
                self.line2d_spec.set_ydata(now_intensities)
                # set max int. line visible if not visible
                self.line2d_spec_int_max.set_visible(True)
                if (self.should_reset_spec_limits == True):
                    # set y-scale
                    self.axes_spectrum.set_ylim(0, 180000)
                    self.should_reset_spec_limits = False
            elif (intensity_type == "Raw Int. - Ref" and self.has_reference_spec):
                logger.debug("Updating y-data for raw - ref")
                # update line data
                calc_intensities = now_intensities - self.reference_spec
                self.line2d_spec.set_ydata(calc_intensities)
                self.line2d_spec_int_max.set_visible(False)
                if (self.should_reset_spec_limits == True):
                    # set y-scale
                    self.axes_spectrum.set_ylim(self.spec_plot_ylims_sub)
                    self.should_reset_spec_limits = False
            elif (intensity_type == "%T or %R" and self.has_reference_spec):
                logger.debug("Updating y-data for %T/%R")
                calc_intensities = (now_intensities - self.dark_spec) / (self.reference_spec - self.dark_spec)
                self.line2d_spec.set_ydata(calc_intensities)
                self.line2d_spec_int_max.set_visible(False)
                if (self.should_reset_spec_limits == True):
                    # set y-scale
                    self.axes_spectrum.set_ylim(self.spec_plot_ylims_t)
                    self.should_reset_spec_limits = False
            elif (intensity_type == "Abs" and self.has_reference_spec):
                logger.debug("Updating y-data for Abs")
                calc_T = (now_intensities - self.dark_spec) / (self.reference_spec - self.dark_spec)
                calc_intensities = np.log10(calc_T)
                self.line2d_spec.set_ydata(calc_intensities)
                self.line2d_spec_int_max.set_visible(False)
                if (self.should_reset_spec_limits == True):
                    # set y-scale
                    self.axes_spectrum.set_ylim(self.spec_plot_ylims_abs)
                    self.should_reset_spec_limits = False
            logger.debug("draw_spec: calling draw_idle()")
            self.canv_spectrum.draw_idle()
            logger.debug("draw_spec: calling flush_events()")
            self.canv_spectrum.flush_events()
        # loop the draw call
        logger.debug("draw_spec: End of call")
        self.root.after(1000, self.draw_spec)
    ''' # Legacy code, TBR
    def draw_spec(self):
        while (self.should_draw_spec):

            if (self.has_spectrometer == False):
                continue
            # get the latest intensities
            now_intensities = self.spectrometer.intensities(self.enable_dark_correction, self.enable_nonlinearity_correction)

            # figure out what kind of spectrum we are drawing
            intensity_type = self.spec_intensity_type.get()
        
            if (intensity_type == "Raw Int."):
                # clear the figure's axes
                self.axes_spectrum.clear()
                # plot the raw intensities
                self.axes_spectrum.plot(self.wavelengths, now_intensities, color="blue")
                # plot a "max intensity" line at 170,000 intensity
                self.axes_spectrum.plot([0, 3000], [170000, 170000], "--", color="red")
                # set plot labels, etc.
                self.axes_spectrum.set_xlabel("Wavelength (nm)")
                self.axes_spectrum.set_ylabel("Raw Intensity (a.u.)")
                self.axes_spectrum.set_xlim(self.wavelengths[0], self.wavelengths[-1])
                self.axes_spectrum.set_ylim(0, 180000)
                self.axes_spectrum.grid()
                # finally execute the draw call
                self.canv_spectrum.draw()
            elif (intensity_type == "Raw Int. - Ref" and self.has_reference_spec):
                calc_intensities = now_intensities - self.reference_spec
                self.axes_spectrum.clear()
                self.axes_spectrum.plot(self.wavelengths, calc_intensities, color="blue")
                self.axes_spectrum.set_xlabel("Wavelength (nm)")
                self.axes_spectrum.set_ylabel("Subtracted Intensity")
                self.axes_spectrum.set_xlim(self.wavelengths[0], self.wavelengths[-1])
                self.axes_spectrum.grid()
                self.canv_spectrum.draw()
            elif (intensity_type == "%T or %R" and self.has_reference_spec):
                calc_intensities = (now_intensities - self.dark_spec) / (self.reference_spec - self.dark_spec)
                self.axes_spectrum.clear()
                self.axes_spectrum.plot(self.wavelengths, calc_intensities, color="blue")
                self.axes_spectrum.set_xlabel("Wavelength (nm)")
                self.axes_spectrum.set_ylabel("Fractional T or R")
                self.axes_spectrum.set_xlim(self.wavelengths[0], self.wavelengths[-1])
                self.axes_spectrum.grid()
                self.canv_spectrum.draw()
            elif (intensity_type == "Abs" and self.has_reference_spec):
                fraction = (now_intensities - self.dark_spec) / (self.reference_spec - self.dark_spec)
                calc_abs = -1 * np.log10(fraction)
                self.axes_spectrum.clear()
                self.axes_spectrum.plot(self.wavelengths, calc_abs)
                self.axes_spectrum.set_xlabel("Wavelength (nm)")
                self.axes_spectrum.set_ylabel("Absorbance")
                self.axes_spectrum.set_xlim(self.wavelengths[0], self.wavelengths[-1])
                self.axes_spectrum.grid()
                self.canv_spectrum.draw()
            
            # sleep until next run
            time.sleep(self.spec_draw_time)
    '''

    # Sets integration time based off the value user has in the box for it
    def set_integ_time(self):
        # only need to do anything if spectrometer is connected
        if (self.has_spectrometer and self.running == False):
            # Try to convert the text to a float, if it doesn't work let the user know
            try:
                integ_time_micros = 1000 * float(self.integ_time_txt.get())
                # Double check that integ time falls within limits
                limits = self.spectrometer.integration_time_micros_limits
                if (integ_time_micros > limits[0]) and (integ_time_micros < limits[1]):
                    with self.spectrometer_lock:
                        self.spectrometer.integration_time_micros(integ_time_micros)
                        self.integration_time_micros = integ_time_micros
                else:
                    mbox.showwarning("Could not set integ. time", "Value is outside hardcoded spectrometer limits")
            except ValueError:
                mbox.showwarning("Could not set integ. time", "Value does not appear to be a string")

    # Stores reference spectrum based on current spectrometer input
    def store_reference_spectrum(self):
        if (self.has_spectrometer and self.running == False):
            with self.spectrometer_lock:
                self.reference_spec = self.spectrometer.intensities(self.enable_dark_correction, self.enable_nonlinearity_correction)
            self.has_reference_spec = True
            logger.info("Stored reference spectrum successfully")
            print("Stored reference spectrum successfully")
    
    # Stores dark spectrum based on current spectrometer input
    # Dark spectrum is subtracted from what we actually measure
    def store_dark_spectrum(self):
        if (self.has_spectrometer and self.running == False):
            with self.spectrometer_lock:
                self.dark_spec = self.spectrometer.intensities(self.enable_dark_correction, self.enable_nonlinearity_correction)
            logger.info("Stored dark spectrum successfully")
            print("Stored dark spectrum successfully")

    # If a spectrometer is attached, collect a spectrum and save right away
    def collect_spec_now(self):
        logger.info("Trying to collect spectrum now...")
        if (self.has_spectrometer == False):
            logger.warning("No spectrometer connected to collect from!")
            return
        
        # Make sure we have what we need to calc. requested intensity type
        if (self.spec_intensity_type.get() != "Raw Int." and self.has_reference_spec == False):
            logger.warning("No reference spectrum saved, but is needed to calculate intensity type")
            mbox.showwarning("Error saving file", "Chosen intensity type requires a reference spectrum")
            return

        # Get the current intensities
        with self.spectrometer_lock:
            now_intensities = self.spectrometer.intensities(self.enable_dark_correction, self.enable_nonlinearity_correction)
        # Get the current time
        now = datetime.datetime.now()

        # Convert the intensity type from the drop down menu to one more filename friendly
        filename_intensity_type = {"Raw Int.": "Raw_Intensity",
                                    "Raw. Int. - Ref": "Reference_Sub_Intensity",
                                    "%T or %R": "Transmittance",
                                    "Abs": "Abs"}

        intensity_type = self.spec_intensity_type.get()
        intensity_type_text = filename_intensity_type[intensity_type]
        # Assemble the filename
        filename = f"{intensity_type_text}_{now.isoformat()}.csv"
        # Replace colons in the filename (windows doesn't like them)
        filename = filename.replace(":", "-")
        # add the folder name
        filename = f"{self.save_dir}/{filename}"
        # Try and save the file
        logger.info("Trying to open file to write current spectrum to")
        with open(filename, "w") as outfile:
            # First write the header
            outfile.write("Ocean Optics spectrometer spectrum generated by Lighthouse\n")
            outfile.write(f"Spectrometer model: {self.spectrometer.model}\n")
            outfile.write(f"Integration time (ms): {self.integration_time_micros/1000:.0f}\n")
            outfile.write(f"Electric dark correction enabled: {self.enable_dark_correction.get()}\n")
            outfile.write(f"Nonlinearity correction enabled: {self.enable_nonlinearity_correction.get()}\n")
            outfile.write(f"Number of data points: {len(self.wavelengths):d}\n\n")
            outfile.write(">>>Begin Data<<<\n")
            outfile.write(f"Wavelength_nm,{self.spec_intensity_type.get()},")
            if (self.has_reference_spec):
                outfile.write("Reference Int.,")
            outfile.write("Dark Int.\n")

            # Then write out all the rows sequentially
            for i in range(0, len(self.wavelengths)):

                # Write the wavelength first
                outfile.write(f"{self.wavelengths[i]},")
                # Then, calculate the desired intensity type, and write that to the file
                if self.spec_intensity_type.get() == "Raw Int.":
                    outfile.write(f"{now_intensities[i]},")
                elif self.spec_intensity_type.get() == "Raw Int. - Ref":
                    out_intensity = now_intensities[i] - self.reference_spec[i]
                    outfile.write(f"{out_intensity},")
                elif self.spec_intensity_type.get() == "%T or %R":
                    out_intensity = (now_intensities[i] - self.dark_spec[i]) / (self.reference_spec[i] - self.dark_spec[i])
                    outfile.write(f"{out_intensity},")
                elif self.spec_intensity_type.get() == "Abs":
                    transmittance = (now_intensities[i] - self.dark_spec[i]) / (self.reference_spec[i] - self.dark_spec[i])
                    out_intensity = -1*np.log10(transmittance)
                    outfile.write(f"{out_intensity},")
                
                # Then write the reference spectrum, if there is one
                if (self.has_reference_spec):
                    outfile.write(f"{self.reference_spec[i]},")
                # Finally write the dark spectrum and end the line
                outfile.write(f"{self.dark_spec[i]}\n")
        logger.info("Successfully wrote file")

    # Function to plot the pstat data from currently running experiment
    def plot_pstat_curve(self):
        # pause very briefly to allow for initial data collection
        logger.info("Starting pstat curve plotting")
        time.sleep(0.3)
        while (self.should_draw_pstat):
            logger.debug("plot_pstat_curve: Top of pstat draw loop")
            now_pstat_pt = self.acq_curve.last_data_point()
            num_pts = self.acq_curve.count()
            # if there aren't enough points, skip plotting for now
            if (num_pts < 2 or now_pstat_pt is None):
                logger.debug("Not enough to pts to plot")
                continue
            
            # get all the currently acquired data
            logger.debug("plot_pstat_curve: calling acq_data()")
            data = self.acq_curve.acq_data()
            # get the elapsed time (to determine when to plot)
            elapsed_time = now_pstat_pt[1]
            # get columns 2 and 4 from the data table, which correspond to potentials in V and current in A respectively

            # try block in here temporarily to see if it gets around this thread crashing on start
            try:
                potentials = data["vf"]
                currents = data["im"]
                cycles = data["cycle"]
            except IndexError:
                print("Index error accessing CV data for plotting")
                continue
            
            # down sample as needed
            if (num_pts < 10000):
                # full # of points if less than 10k
                logger.debug("plot_pstat_curve: Plotting all points in pstat curve")
                plot_potentials = potentials
                plot_currents = currents
                plot_cycles = cycles
            elif (num_pts  < 100000):
                # every 5th point if btwn 10k-100k pts
                logger.debug("plot_pstat_curve: Plotting every 5th point in pstat curve")
                plot_potentials = potentials[::5]
                plot_currents = currents[::5]
                plot_cycles = cycles[::5]
            else:
                # if we have > 100,000 pts, only plot every 10th point to save on memory
                logger.debug("plot_pstat_curve: Plotting every 10th point in pstat curve")
                plot_potentials = potentials[::10]
                plot_currents = currents[::10]
                plot_cycles = cycles[::10]

            # clear the old plot
            logger.debug("plot_pstat_curve: calling clear()")
            self.axes_cv.clear()
            # find the first position of the current cycle
            now_cycle = now_pstat_pt[8]
            now_cycle_index = np.argmax(plot_cycles == now_cycle) # this returns 0 if the first index is 0 or 0 on a fail to find
            if (now_cycle_index == 0):
                # in the "0 index case" - we have either only the first cycle, 
                # or some weird error where we couldn't find the current cycle in our data here
                # regardless, just plot everything as-is
                logger.debug("plot_pstat_curve: calling plot() for all cycles in blue")
                self.axes_cv.plot(plot_potentials, plot_currents, color="blue")
            else: # if we have > 1 cycle & we can find an index, sketch current cycle in different color
                # plot everything up to where the current cycle starts in blue
                logger.debug("plot_pstat_curve: calling plot() for previous cycles in blue")
                self.axes_cv.plot(plot_potentials[:now_cycle_index], plot_currents[:now_cycle_index], color="blue")
                # plot everything after in red
                logger.debug("plot_pstat_curve: calling plot() for current cycle in red")
                self.axes_cv.plot(plot_potentials[now_cycle_index:], plot_currents[now_cycle_index:], color="red")
            
            # label axes
            self.axes_cv.set_xlabel("WE Potential (V)")
            self.axes_cv.set_ylabel("Current (A)")
            # grid
            self.axes_cv.grid()
            # draw the updates
            logger.debug("plot_pstat_curve: calling draw()")
            self.canv_cv.draw()
            logger.debug("Bottom of pstat draw loop")
            # if the time is past the first minute, plot only every 30 s
            if (elapsed_time > 30):
                time.sleep(30)
            else:
                time.sleep(3)

    def start_measurement(self):
        logger.info("Attempting to start measurement")
        # if already running, do nothing
        if (self.running):
            logger.warning("Aborting: Already running a measurement")
            return
        ### First, check if experiment settings are valid

        # Check PStat parameters here
        try:
            v1_num = float(self.vertex_pot1_text.get())
            v2_num = float(self.vertex_pot2_text.get())
            scanrate_num = float(self.scan_rate_text.get())
            cycles_num = int(self.num_cycles_text.get())
            sample_period = float(self.step_size_text.get())
        except ValueError:
            logger.error("Aborting: One of the PStat parameters does not appear to be a valid number.")
            mbox.showwarning("Value Error", "One of the PStat parameters does not appear to be a valid number.")
            return

        self.current_expt_max_cycles = cycles_num

        # Make sure collection time > integration time
        integration_time_ms = self.integration_time_micros / 1000
        # try to get collection time, give an error if the text is not a valid number
        try:
            collection_time_num = float(self.spec_freq_txt.get())
        except ValueError:
            logger.error("Aborting: Could not convert contents of collection freq. box into a number")
            mbox.showwarning("Value Error", "Could not convert contents of collection frequency box into a number!")
            return
        # Convert collection time into ms for comparison
        collection_time_unit = self.spec_freq_units.get()
        if (collection_time_unit == "ms"):
            collection_time_ms = collection_time_num * 1
        elif (collection_time_unit == "s"):
            collection_time_ms = collection_time_num * 1000
        elif (collection_time_unit == "min"):
            collection_time_ms = collection_time_num * 60*1000
        elif (collection_time_unit == "hr"):
            collection_time_ms = collection_time_num * 3600*1000
        
        # Finally verify that collection time is indeed bigger than integ time
        if (collection_time_ms <= integration_time_ms):
            logger.error("Aborting: collection frequency is <= integration time")
            mbox.showwarning("Collection Frequency Error", "Minimum time between collecting spectra must be larger than integration time!")
            return
        
        ### Next - make sure we have what we need to calculate the requested intensity type
        if (self.spec_intensity_type.get() != "Raw Int." and (self.has_reference_spec == False)):
            logger.error("Aborting: no reference spectrum and intensity type requires it.")
            mbox.showwarning("Error saving file", "Chosen intensity type requires a reference spectrum")
            return
        
        ### check if the potentiostat and spectrometer are connected, if they are needed for this experiment
        if (self.use_pstat.get() and self.has_potentiostat == False):
            mbox.showwarning("Potentiostat Not Connected", "You are using the potentiostat for this measurement, but it does not appear to be connected.")
            return
        if (self.use_spec.get() and self.has_spectrometer == False):
            mbox.showwarning("Spectrometer Not Connected", "You are using the spectrometer for this measurement, but it does not appear to be connected.")
            return
        
        ### Next - try to open a file with experiment name and date

        # Convert the intensity type from the drop down menu to one more filename friendly
        filename_intensity_type = {"Raw Int.": "Raw_Intensity",
                                    "Raw Int. - Ref": "Reference_Sub_Intensity",
                                    "%T or %R": "Transmittance",
                                    "Abs": "Abs"}
        self.running_intensity_type = self.spec_intensity_type.get() # Make sure we always acquire w/ the same intensity type
        intensity_type_text = filename_intensity_type[self.spec_intensity_type.get()]
        # Assemble the filename
        now = datetime.datetime.now() # Get the current time for the name
        self.filename_time = now
        filename = f"{self.experiment_name}_{now.isoformat()}_{intensity_type_text}.csv"
        filename = filename.replace(":", "-")
        # Append directory
        filename = f"{self.save_dir}/{filename}"
        # Make sure the output file is opened correctly
        logger.info(f"Trying to open main expt data file w/ name: {filename}")
        try:
            self.outfile = open(filename, mode="w")
        except OSError as e:
            logger.error(f"Error opening file: {e.strerror}")
            mbox.showerror("Error opening file", e.strerror)
            return
        # save the filename to class for zipping later
        self.filename = filename

        ### Next - attempt to open and write to files for the reference and dark spectra

        # If we have a reference spectrum, try to write it to a file
        if (self.has_reference_spec):
            reference_filename = f"{self.experiment_name}_{now.isoformat()}_REFERENCE_SPECTRUM.csv"
            reference_filename = reference_filename.replace(":", "-")
            reference_filename = f"{self.save_dir}/{reference_filename}"
            logger.info(f"Trying to write reference spectrum to disk @ filename: {reference_filename}")
            with open(reference_filename, "w") as reference_file:
                reference_file.write("Ocean Optics spectrometer reference spectrum generated by Lighthouse\n")
                reference_file.write(f"For experiment: {self.experiment_name}\n")
                reference_file.write(f"Experiment began at: {now.isoformat()}\n")
                reference_file.write(f"Spectrometer model: {self.spectrometer.model}\n")
                reference_file.write(f"Integration time (ms): {self.integration_time_micros/1000:.0f}\n")
                reference_file.write(f"Electric dark correction enabled: {self.enable_dark_correction.get()}\n")
                reference_file.write(f"Nonlinearity correction enabled: {self.enable_nonlinearity_correction.get()}\n")
                reference_file.write(f"Number of data points: {len(self.wavelengths):d}\n\n")
                reference_file.write(">>>Begin Data<<<\n")
                reference_file.write("Wavelength_nm,Intensity")
                # then write all the data points
                for i in range(0, len(self.wavelengths)):
                    wavelength = self.wavelengths[i]
                    intensity = self.reference_spec[i]
                    reference_file.write(f"{wavelength},{intensity}\n")
                # flush and close
                reference_file.close()
                # save reference filename for zipping later
                self.reference_filename = reference_filename
                logger.info("Succeeded in writing reference spectrum")
        
        # Try to write the dark spectrum
        dark_filename = f"{self.experiment_name}_{now.isoformat()}_DARK_SPECTRUM.csv"
        dark_filename = dark_filename.replace(":", "-")
        dark_filename = f"{self.save_dir}/{dark_filename}"
        logger.info("Trying to write dark spectrum to disk...")
        with open(dark_filename, "w") as dark_file:
            dark_file.write("Ocean Optics spectrometer dark spectrum generated by Lighthouse\n")
            dark_file.write(f"For experiment: {self.experiment_name}\n")
            dark_file.write(f"Experiment began at: {now.isoformat()}\n")
            dark_file.write(f"Spectrometer model: {self.spectrometer.model}\n")
            dark_file.write(f"Integration time (ms): {self.integration_time_micros/1000:.0f}\n")
            dark_file.write(f"Electric dark correction enabled: {self.enable_dark_correction.get()}\n")
            dark_file.write(f"Nonlinearity correction enabled: {self.enable_nonlinearity_correction.get()}\n")
            dark_file.write(f"Number of data points: {len(self.wavelengths):d}\n\n")
            dark_file.write(">>>Begin Data<<<\n")
            dark_file.write("Wavelength_nm,Intensity")
            # write all the data points
            for i in range(0, len(self.wavelengths)):
                wavelength = self.wavelengths[i]
                intensity = self.dark_spec[i]
                dark_file.write(f"{wavelength},{intensity}\n")
            dark_file.close()
            # save dark filename for zipping later
            self.dark_filename = dark_filename
            logger.info("Succeeded in writing dark spectrum")
        
        logger.info("Writing main data file header")
        ### Next - write the header for our main data file
        self.outfile.write("Spectroelectrochemistry data file generated by Lighthouse\n")
        self.outfile.write(f"For experiment: {self.experiment_name}\n")
        self.outfile.write(f"Operator: {self.operator}\n")
        self.outfile.write(f"Description: {self.description}\n")
        self.outfile.write(f"Experiment began at: {now.isoformat()}\n")
        # include potentiostat info if we are using it
        self.outfile.write(f"Use Potentiostat?: {self.use_pstat.get()}\n")
        if (self.use_pstat.get()):
            self.outfile.write(f"Potentiostat Model Number: {self.potentiostat.model_no()}\n")
            self.outfile.write(f"Vertex Potential 1 (V): {self.vertex_pot1_text.get()}\n")
            self.outfile.write(f"Vertex Potential 2 (V): {self.vertex_pot2_text.get()}\n")
            self.outfile.write(f"Scan Rate (V/s): {self.scan_rate_text.get()}\n")
            self.outfile.write(f"Sample Time (s): {self.step_size_text.get()}\n")
            self.outfile.write(f"Max # Cycles: {self.num_cycles_text.get()}\n")
        # include spectrometer info if we are using it
        self.outfile.write(f"Use Spectrometer?: {self.use_spec.get()}\n")
        if (self.use_spec.get()):
            self.outfile.write(f"Spectrometer model: {self.spectrometer.model}\n")
            self.outfile.write(f"Integration time (ms): {self.integration_time_micros/1000:.0f}\n")
            self.outfile.write(f"Electric dark correction enabled: {self.enable_dark_correction.get()}\n")
            self.outfile.write(f"Nonlinearity correction enabled: {self.enable_nonlinearity_correction.get()}\n")
            self.outfile.write(f"Intensity Type: {self.running_intensity_type}\n")
            self.outfile.write(f"Number of distinct wavelengths: {len(self.wavelengths):d}\n")
        self.outfile.write("Column headers to the right of potentiostat data are the spectrometer wavelengths in nm.\n\n")
        self.outfile.write(">>>Begin Data<<<\n")
        self.outfile.write("Time_s,Cycle_num,Potential_V,Current_A")

        # write the wavelengths in the column headers
        for wavelength in self.wavelengths:
            self.outfile.write(f",{wavelength:.1f}")

        ### Next - Set up the voltage waveform and data saving object


        # NOTE: Gamry sets a maximum number of points in this signal to 2^18-1, or ~260,000 pts. 
        # If you input a sample period that puts the # of points > ~260,000, then the signal creation
        # here will throw an error and the experiment won't run. 
        # For 200 cycles, scanning from +0 to +1 V, at 10 mV/s, this means the highest sample period you can have should be ~0.2s
        logger.info("Attempting to set up pstat singal. May throw an error if # of pts would be too high")
        self.ramp_signal = self.potentiostat.signal_r_up_dn_new([v1_num, v1_num, v2_num, v2_num], 
                                                       [scanrate_num, scanrate_num, scanrate_num], 
                                                       [0, 0, 0], sample_period, cycles_num, tkp.PSTATMODE)
        
        # Set signal for potentiostat to this new signal
        self.potentiostat.set_signal_r_up_dn(self.ramp_signal)
        logger.info("Attempting to initialize pstat signal")
        self.potentiostat.init_signal()
        logger.info("Succeeded in initializing pstat signal")
        
        # Initialize the data collection curve
        # the second number here is the max number of data points in the buffer
        # if the number of points exceeds this number during the run, the oldest points are overwritten
        # if we continuously output to file this shouldn't be an issue as long as we aren't collecting data
        # at an ultra fast rate (> a few hundred ms)

        # The support people at Gamry tell me that the max range is somewhere a little above 5 million points
        # With 4 million points, scanning a 1V potential window with a 1 mV step gives us a max of ~2000 cycles
        logger.info("Attempting to initialize pstat acq curve")
        self.acq_curve = tkp.RcvCurve(self.potentiostat, 4000000)
        self.acq_curve.set_stop_i_max(True, 5) # automatically stop if current exceeds 5 A

        ### Next - show that the exp is running on the GUI
        # self.lbl_running.configure(text=f"Running: Cycle 0/{self.current_expt_max_cycles}", fg="green") TBR

        ### Next - Actually begin the run!
        self.was_aborted = False # flag for if expt was forcibly ended prematurely
        self.time_start = time.time()
        self.calc_run_time = np.abs(v2_num-v1_num) / scanrate_num * 2 * cycles_num
        print(f"Estimated experiment time: {self.calc_run_time/60:.2f} minutes.")
        self.num_freq_s = collection_time_ms/1000
        logger.info("Constructing measurement thread")
        self.thread_measurement = threading.Thread(target=self.run_measurement)
        self.thread_measurement.start()

    # The actual looping action of measuring
    def run_measurement(self):
        logger.info("Start of measurement thread")
        # the timestamp of the last flush to disk in s
        self.last_file_flush = -100
        # control flags
        self.running = True
        self.should_draw_pstat = True
        # thread where the potentiostat curve is drawn
        self.thread_draw_pstat = threading.Thread(target=self.plot_pstat_curve)
        # Turn on the cell
        #self.lbl_cell_state.configure(text="Cell On", fg="green") TBR
        logger.info("Setting cell ON")
        self.potentiostat.set_cell(True)
        # Start running the data acquisition curve
        logger.info("Starting data acq.")
        self.acq_curve.run(True)
        self.thread_draw_pstat.start()

        # # of data points counter
        i = 0
        # loop start time
        loop_start_time = time.perf_counter_ns()
        while self.acq_curve.running():
            i += 1
            logger.debug(f"run_measurement: Top of loop; i={i}")
            # Grab the most recent data point from PStat, including time, V, i, etc.
            now_pstat_pt = self.acq_curve.last_data_point()
            self.most_recent_pstat_pt = now_pstat_pt
            elapsed_time = now_pstat_pt[1]
            now_potential = now_pstat_pt[2]
            now_current = now_pstat_pt[4]
            now_cycle = now_pstat_pt[8]
            # Grab the most recent spectrum, blocking spectrometer access until it returns
            logger.debug("run_measurement: Trying to get spectrometer intensities")
            with self.spectrometer_lock:
                now_raw_intensities = self.spectrometer.intensities(self.enable_dark_correction, self.enable_nonlinearity_correction)
            logger.debug("run_measurement: Beginning calculating intensities")
            # calculate desired intensity type
            if (self.running_intensity_type == "Raw Int."):
                now_intensities = now_raw_intensities
            elif (self.running_intensity_type == "Raw Int. - Ref"):
                now_intensities = now_raw_intensities - self.reference_spec
            elif (self.running_intensity_type == "%T or %R"):
                now_intensities = (now_raw_intensities - self.dark_spec) / (self.reference_spec - self.dark_spec)
            else:
                # calc as abs
                now_intensities = (now_raw_intensities - self.dark_spec) / (self.reference_spec - self.dark_spec)
                now_intensities = -1*np.log10(now_intensities)

            logger.debug("run_measurement: Begin writing this row to disk")
            # Output this row to file
            self.outfile.write("\n")
            # Write time, potential, current
            self.outfile.write(f"{elapsed_time},{now_cycle},{now_potential},{now_current}")
            # Write intensities
            for intensity in now_intensities:
                self.outfile.write(f",{intensity}")
            logger.debug("run_measurement: Row written to disk")
            

            # flush to disk every 30 s
            if (elapsed_time - self.last_file_flush > 30):
                logger.debug("run_measurement: flushing file to disk")
                self.outfile.flush()
                self.last_file_flush = elapsed_time

            # update # of cycles label
            #self.lbl_running.configure(text=f"Running: Cycle {now_cycle}/{self.current_expt_max_cycles}", fg="green") TBR
            # calculate the time when we need to take the next data point
            logger.debug("run_measurement: begin calculating next time we need to wait for")
            try:
                # note - the cast to int here is necessary because ints in python can grow arbitrarily large
                # floats can overflow. This number overflowing can cause hangs that interrupt a measurement
                next_pt_time = np.floor(loop_start_time + i * int(self.num_freq_s) * (10**9))
            except FloatingPointError:
                logger.error(f"run_measurement: Integer overflow or underflow detected. Terminating measurement.")
                self.abort_measurement()
            except OverflowError:
                logger.error(f"run_measurement: Integer overflow or underflow detected. Terminating measurement.")
                self.abort_measurement()
            if (math.isinf(next_pt_time)):
                logger.error(f"Overflow detected. Aborting measurement")
                self.abort_measurement()
            logger.debug("run_measurement: begin sleep til next time")
            perf_sleep_until(next_pt_time)
            # Below is legacy code - TBR
            # loop_end_time = time.perf_counter_ns()
            # dt_s = (loop_end_time - loop_start_time) / (10**9)
            # sleep_time = self.num_freq_s - dt_s
            # # pause til next step
            # if (sleep_time > 0.015):
            #     time.sleep(self.num_freq_s)
        
        # Finish the measurement
        # Turn off the cell
        logger.info("run_measurement: Setting cell OFF")
        self.potentiostat.set_cell(False)
        #self.lbl_cell_state.configure(text="Cell Off", fg="red") TBR
        # Close the file handle and flush to disk
        logger.info("run_measurement: Closing and flushing main datafile to disk")
        self.outfile.close()
        # stop drawing pstat curve
        self.should_draw_pstat = False
        # get end time
        self.end_time = time.time()
        # Finally, write the acquisition curve to disk
        # The main CSV file has already been outputted, but this file, handled by the Gamry library,
        # Contains all the potentiostat data points, even the ones in between the ones in the CSV file
        # This file will be most useful if one wants to plot the CV as standalone data
        now = self.filename_time
        pstat_data_filename = f"{self.experiment_name}_Raw_PStat_Data_{now}.csv"
        pstat_data_filename = pstat_data_filename.replace(":", "-")
        pstat_data_filename = f"{self.save_dir}/{pstat_data_filename}"
        logger.info("run_measurement: Attempting to write raw pstat data file to disk...")
        with open(pstat_data_filename, "w") as file_pstatdata:
            '''Side note here: Instead of pretty formatting as I've done for the files everywhere else
               in this program, I'm just outputting the data table as Gamry stores it.
               This should be the same as their regular output DTA files but I am not sure.

               If I need to I can come back and clean this code up, but since it's a niche feature,
               I thought it would be fine like this.
            '''
            np.savetxt(file_pstatdata, self.acq_curve.acq_data(), delimiter=",", newline="\n",
            header="point,time,vf,vu,im,ach,vsig,temp,cycle,ie_range,overload,stop_test", comments='')
            # Save raw pstat data filename for zipping later 
            self.pstat_data_filename = pstat_data_filename
            logger.info("run_measurement: Done!")
        # update GUI
        #self.lbl_running.configure(text="not running", fg="red") TBR
        self.running = False
        # try to send emails out to notify experiment is complete if the emails are there and expt wasn't ended early on purpose
        if (self.emails != "" and self.was_aborted == False):
            logger.info("Trying to send emails...")
            self.try_send_notif_emails()
            self.try_send_file_emails()
    
    def try_send_notif_emails(self):
        # Note: Sender email and password are in auth_token.txt file.
        port = 465  
        try:
            authfile = open("auth_token.txt")
        except:
            print("Error opening file containing the auth token")
            return
        sender_email = authfile.readline().strip()
        password = authfile.readline().strip()
        # Parse recevier emails
        receiver_emails = self.emails.split(",")

        # Make a new MIME message
        message = MIMEMultipart()
        message["From"] = "Lighthouse Spectroelectrochemistry Program"
        message["To"] = self.emails
        message["Subject"] = "Lighthouse Measurement Ended"

        # Get end time of expt and convert to readable format
        end_timestamp = datetime.datetime.fromtimestamp(self.end_time)
        end_time_string = end_timestamp.strftime("%a, %Y/%m/%d, %I:%M:%S %p")

        body = """This is an automated message from the Graham lab potentiostat/Ocean Optics spectroelectrochemistry setup.
          This email is to notify you that the experiment concluded (successfully or unsuccesfully) at """
        body = body + end_time_string + "."
        
        # attach the body
        message.attach(MIMEText(body, "plain"))
        # data could also be attached here later

        # Send the message
        msg_text = message.as_string()
        context = ssl.create_default_context()
        try:
            server = smtplib.SMTP_SSL("smtp.gmail.com", port, context=context)
            server.login(sender_email, password)
            server.sendmail(sender_email, receiver_emails, msg_text)
        except smtplib.SMTPConnectError:
            print("Could not connect to SMTP server; this may be due to an internet issue on your side or GMail outage")
        except smtplib.SMTPAuthenticationError:
            print("Error authenticating the automatic sender address. Most likely the login token is expired. Make a new email address and token or contact Carter Pryor.")


    def try_send_file_emails(self):
        # first, try to create the zip file
        zip_path = self.filename[:-4] + ".zip"
        try:
            myzip = zipfile.ZipFile(zip_path, mode="w", compression=zipfile.ZIP_DEFLATED, compresslevel=9)
            # write each of the files to the zip archive
            rel_filename = self.filename.split("/")[-1]
            myzip.write(self.filename, arcname=rel_filename)

            rel_dark_filename = self.dark_filename.split("/")[-1]
            myzip.write(self.dark_filename, arcname=rel_dark_filename)

            rel_pstat_data_filename = self.pstat_data_filename.split("/")[-1]
            myzip.write(self.pstat_data_filename, arcname=rel_pstat_data_filename)
            if (self.has_reference_spec):
                rel_ref_spec_filename = self.reference_filename.split("/")[-1]
                myzip.write(self.reference_filename, arcname=rel_ref_spec_filename)
            # close and flush file to disk
            myzip.close()
        except:
            # if something goes wrong, just let user know.
            print("Failed to create zip file to send to emails")
            return

        # then construct and send the email
        # Note: Sender email and password stored in auth_token.txt file
        port = 465
        try:
            authfile = open("auth_token.txt")
        except:
            print("Error opening file containing the auth token")
            return
        sender_email = authfile.readline().strip()
        password = authfile.readline().strip()
        # Parse recevier emails
        receiver_emails = self.emails.split(",")

        # Make a new MIME message
        message = MIMEMultipart()
        message["From"] = "Lighthouse Spectroelectrochemistry Program"
        message["To"] = self.emails
        message["Subject"] = "Lighthouse Measurement Data"

        # Get end time of expt and convert to readable format
        end_timestamp = datetime.datetime.fromtimestamp(self.end_time)
        end_time_string = end_timestamp.strftime("%a, %Y/%m/%d, %I:%M:%S %p")

        body = """This is an automated message from the Graham lab potentiostat/Ocean Optics spectroelectrochemistry setup.
          This email contains the data for the experiment that concluded at """
        body = body + end_time_string + "."
        
        # attach the body
        message.attach(MIMEText(body, "plain"))
        # try to attach data too
        try:
            attachment_file = open(zip_path, 'rb')
            # Add file as "application/zip" (MIME subtype indicating generic binary file)
            part = MIMEBase("application", "zip")
            part.set_payload(attachment_file.read())
            # Encode binary file into ASCII to be able to send via email
            encoders.encode_base64(part)
            # Add a header as a key/value pair to the attachment
            zip_path_parts = zip_path.split("/")
            zip_name = zip_path_parts[-1]
            part.add_header(
                "Content-Disposition",
                f"attachment; filename={zip_name}",
            )
            # then attach the file to the email
            message.attach(part)
        except:
            # error attaching zip file to email
            print("Error opening zip file or attaching to email")
            return


        # Send the message
        msg_text = message.as_string()
        context = ssl.create_default_context()
        try:
            server = smtplib.SMTP_SSL("smtp.gmail.com", port, context=context)
            server.login(sender_email, password)
            server.sendmail(sender_email, receiver_emails, msg_text)
        except smtplib.SMTPConnectError:
            print("Could not connect to SMTP server; this may be due to an internet issue on your side or GMail outage")
        except smtplib.SMTPAuthenticationError:
            print("Error authenticating the automatic sender address. Most likely the login token is expired. Make a new email address and token or contact Carter Pryor.")
      

    def abort_measurement(self):
        # only do anything if the acquisition is actually running
        if (self.running):
            # After "stop" is called, the measurement thread should automatically terminate its while loop
            # which should write the data to disk and finish cleanup after measurement
            self.was_aborted = True
            self.acq_curve.stop()


    # Destructor - make sure to close the potentiostat so other programs can use it when we are done       
    def __del__(self): 
        if (self.has_potentiostat):
            self.potentiostat.close()
            # Gamry recommends calling del on all resources even if the program is going to terminate
            del self.ramp_signal
            del self.acq_curve
            del self.potentiostat

# raise exceptions in case of floating point going over/under
np.seterr(over="raise", under="raise")
# Basic configure the logger
# get current time for file name
t_now = datetime.datetime.now()
t_str = t_now.isoformat().replace(":", "-")
# create logs folder if it does not exist
log_folderpath = pathlib.Path("logs")
log_folderpath.mkdir(exist_ok=True)
# begin logging
logging.basicConfig(filename=f"logs/{t_str}.log", level=logging.NOTSET) # NOTSET here means everything will be logged, can change this for proper releases
# Initialize a Window object        
window = MyWindow()
# Schedule GUI update function
window.root.after(100, window.gui_update)
# Run its main GUI loop
window.root.mainloop()
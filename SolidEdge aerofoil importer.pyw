import os
import re
import time
import math
import tkinter as tk
from tkinter import ttk, filedialog, messagebox
import win32com.client
from pywintypes import com_error

_DELAY = 0.05  # seconds
_TIMEOUT = 5.0  # seconds

digit_pattern = r"-?\d+\.\d{6}"
aerofoil_match = re.compile(fr"^\s*({digit_pattern})\s*({digit_pattern})\s*$", flags = re.MULTILINE)
float_match = re.compile(r"^([+-]?)\d*(?:[.,]\d*)?$")

named_planes = {0: "X/Y", 1: "Y/Z", 2: "X/Z"}


def _com_call_wrapper(f, *args, **kwargs):
    """
    COMWrapper support function.
    Repeats calls when 'Call was rejected by callee.' exception occurs.
    """
    # Unwrap inputs
    args = [arg.wrapped_object if isinstance(arg, COMWrapper) else arg for arg in args]
    kwargs = {key: value.wrapped_object if isinstance(value, COMWrapper) else value for key, value in kwargs.items()}

    result = None
    start_time = time.time()
    while True:
        try:
            result = f(*args, **kwargs)
        except com_error as e:
            if e.hresult == -2147418111:
                print("Call was rejected by callee, retrying...")
                if time.time() - start_time >= _TIMEOUT:
                    raise
                time.sleep(_DELAY)
                continue
            raise
        break

    # if isinstance(result, (win32com.client.CDispatch, win32com.client.CoClassBaseClass)) or callable(result):
    if "win32com" in getattr(result, "__module__", "") or callable(result):
        return COMWrapper(result)
    return result


class COMWrapper:
    """
    Class to wrap COM objects to repeat calls when 'Call was rejected by callee.' exception occurs.
    """

    def __init__(self, wrapped_object):
        # assert isinstance(wrapped_object, win32com.client.CDispatch) or callable(wrapped_object)
        self.__dict__['wrapped_object'] = wrapped_object

    def __getattr__(self, item):
        # return _com_call_wrapper(self.wrapped_object.__getattr__, item)
        return _com_call_wrapper(getattr, self, item)

    def __getitem__(self, item):
        return _com_call_wrapper(self.wrapped_object.__getitem__, item)

    def __setattr__(self, key, value):
        # _com_call_wrapper(self.wrapped_object.__setattr__, key, value)
        _com_call_wrapper(setattr, self, key, value)

    def __setitem__(self, key, value):
        _com_call_wrapper(self.wrapped_object.__setitem__, key, value)

    def __call__(self, *args, **kwargs):
        return _com_call_wrapper(self.wrapped_object.__call__, *args, **kwargs)

    def __repr__(self):
        return 'ComWrapper<{}>'.format(repr(self.wrapped_object))

    def __eq__(self, other):
        if isinstance(other, COMWrapper):
            return self.wrapped_object == other.wrapped_object
        return False


class FloatEntry(ttk.Entry):
    """Entry allowing insertion of only float values"""

    def __init__(self, master, default_value = 0, **kwargs):
        super().__init__(master, **kwargs)

        vcm = (self.register(self.validate_number), "%P")
        self.config(validate = "key", validatecommand = vcm)

        self.insert("end", str(default_value))

    @staticmethod
    def validate_number(after: str) -> bool:
        """Validate if given number is float"""
        match = float_match.match(after)
        return match is not None

    def get(self) -> float:
        """Get current value as an int"""
        value = super().get().lstrip("+").replace(",", ".")
        if not value or value in (".", "-.", "+."):
            return 0
        return float(value)


class MainApplication(ttk.Frame):
    def __init__(self, master, app, constants, **kwargs):
        super().__init__(master, **kwargs)

        self.constants = constants
        self.app = app
        self.doc = None
        self.planes = {}
        self.sketches = {}
        self.aerofoil: list = []

        # Selector variables
        self.v_planes = tk.StringVar()
        self.v_sketches = tk.StringVar()

        # Containers
        self.f_buttons = ttk.Frame(self)

        self.f_modify_groups = ttk.Frame(self)
        self.f_move = ttk.LabelFrame(self.f_modify_groups, text = "Move [mm]")
        self.f_scale = ttk.LabelFrame(self.f_modify_groups, text = "Size [mm]")
        self.f_rotate = ttk.LabelFrame(self.f_modify_groups, text = "Rotate [°]")
        self.f_mirror = ttk.LabelFrame(self.f_modify_groups, text = "Mirror")
        self.f_planes = ttk.LabelFrame(self, text = "Plane")
        self.f_sketches = ttk.LabelFrame(self, text = "Sketch")

        # Configure containers
        self.f_move.columnconfigure(0, weight = 1)
        self.f_scale.columnconfigure(1, weight = 1)

        # "Working" widgets
        self.v_placement = tk.StringVar(value = "new")
        self.rb_new = ttk.Radiobutton(self, text = "New sketch", variable = self.v_placement, value = "new")
        self.rb_existing = ttk.Radiobutton(self, text = "Existing sketch", variable = self.v_placement,
                                           value = "existing")
        self.rb_current = ttk.Radiobutton(self, text = "Current sketch", variable = self.v_placement, value = "current")

        self.b_select = ttk.Button(self.f_buttons, text = "Select aerofoil", command = self.load_aerofoil)
        self.b_import = ttk.Button(self.f_buttons, text = "Import", command = self.import_into_se)
        self.l_selected = tk.Label(self.f_buttons, font = ("", 9, "italic"), fg = "grey40")

        self.l_move_x = ttk.Label(self.f_move, text = "X:")
        self.l_move_y = ttk.Label(self.f_move, text = "Y:")
        self.e_move_x = FloatEntry(self.f_move, width = 9, justify = "center")
        self.e_move_y = FloatEntry(self.f_move, width = 9, justify = "center")

        self.v_scale = tk.StringVar(value = "none")
        self.rb_scale_none = ttk.Radiobutton(self.f_scale, text = "Default", variable = self.v_scale, value = "none")
        self.rb_scale_x = ttk.Radiobutton(self.f_scale, text = "W:", variable = self.v_scale, value = "width")
        self.rb_scale_y = ttk.Radiobutton(self.f_scale, text = "H:", variable = self.v_scale, value = "height")
        self.e_scale_x = FloatEntry(self.f_scale, width = 9, justify = "center")
        self.e_scale_y = FloatEntry(self.f_scale, width = 9, justify = "center")

        self.l_rotate = ttk.Label(self.f_rotate, text = "\u27f3:")
        self.e_rotate = FloatEntry(self.f_rotate, width = 9, justify = "center")

        self.v_mirror = tk.StringVar(value = "none")
        self.rb_mirror_none = ttk.Radiobutton(self.f_mirror, text = "None", variable = self.v_mirror, value = "none")
        self.rb_mirror_horizontal = ttk.Radiobutton(self.f_mirror, text = "Horizontally", variable = self.v_mirror,
                                                    value = "horizontal")
        self.rb_mirror_vertical = ttk.Radiobutton(self.f_mirror, text = "Vertically", variable = self.v_mirror,
                                                  value = "vertical")

        # Trace selected radio buttons
        self.v_placement.trace_add("write", self.on_placement_change)
        self.v_scale.trace_add("write", self.on_scale_change)

        # Init other parts of the window
        self.layout_widgets()
        self.on_scale_change()

        # Bindings
        self.focus_set()  # Grab focus so that FocusIn is called, else it would he noticed only by the root window
        self.bind("<FocusIn>", self.reload_se)

    def layout_widgets(self) -> None:
        """Place widgets into application window"""
        self.f_buttons.grid(row = 0, column = 0, columnspan = 3, sticky = "ew", padx = 6, pady = (9, 3))
        self.b_select.pack(side = "left", padx = 3, ipadx = 5)
        self.l_selected.pack(side = "left", padx = 3)
        self.b_import.pack(side = "right", padx = 3, ipadx = 5)

        self.rb_current.grid(row = 1, column = 0, sticky = "w", padx = (9, 3), pady = 3)
        self.rb_new.grid(row = 1, column = 1, sticky = "w", padx = 3, pady = 3)
        self.rb_existing.grid(row = 1, column = 2, sticky = "w", padx = (3, 9), pady = 3)

        self.f_modify_groups.grid(row = 2, column = 0, sticky = "new", padx = (9, 3), pady = (1, 7))
        self.f_planes.grid(row = 2, column = 1, sticky = "new", padx = 3, pady = 3)
        self.f_sketches.grid(row = 2, column = 2, sticky = "new", padx = (3, 9), pady = 3)

        self.f_move.pack(fill = "x", pady = 2)
        self.f_scale.pack(fill = "x", pady = 2)
        self.f_rotate.pack(fill = "x", pady = 2)
        self.f_mirror.pack(fill = "x", pady = 2)

        self.l_move_x.grid(row = 0, column = 0, sticky = "w", padx = (4, 2))
        self.l_move_y.grid(row = 1, column = 0, sticky = "w", padx = (4, 2))
        self.e_move_x.grid(row = 0, column = 1, sticky = "e", padx = (2, 4))
        self.e_move_y.grid(row = 1, column = 1, sticky = "e", padx = (2, 4), pady = 4)

        self.rb_scale_none.grid(row = 0, column = 0, columnspan = 2, sticky = "w", padx = (4, 2))
        self.rb_scale_x.grid(row = 1, column = 0, sticky = "w", padx = (4, 2), pady = (4, 0))
        self.rb_scale_y.grid(row = 2, column = 0, sticky = "w", padx = (4, 2))
        self.e_scale_x.grid(row = 1, column = 1, sticky = "e", padx = (2, 4), pady = (4, 0))
        self.e_scale_y.grid(row = 2, column = 1, sticky = "e", padx = (2, 4), pady = 4)

        self.l_rotate.pack(side = "left", padx = (4, 2), pady = (0, 4))
        self.e_rotate.pack(side = "right", padx = (2, 4), pady = (0, 4))

        self.rb_mirror_none.pack(anchor = "w", padx = 4)
        self.rb_mirror_horizontal.pack(anchor = "w", padx = 4)
        self.rb_mirror_vertical.pack(anchor = "w", padx = 4)

    def on_placement_change(self, *_) -> None:
        """When selected placement option changes, disable other widgets"""
        plane_state = "normal" if self.v_placement.get() == "new" else "disabled"
        for widget in self.f_planes.winfo_children():
            widget.config(state = plane_state)

        sketch_state = "normal" if self.v_placement.get() == "existing" else "disabled"
        for widget in self.f_sketches.winfo_children():
            widget.config(state = sketch_state)

    def on_scale_change(self, *_) -> None:
        """Enable entry widget cor currently selected scale option"""
        self.e_scale_x.config(state = "normal" if self.v_scale.get() == "width" else "disabled")
        self.e_scale_y.config(state = "normal" if self.v_scale.get() == "height" else "disabled")

    def reload_se(self, *_) -> None:
        """Reload SolidEdge connection. Get new active document and load current planes and sketches"""
        # Clear everything
        self.doc = None
        self.planes.clear()
        self.sketches.clear()
        for widget in self.f_planes.winfo_children():
            widget.destroy()
        for widget in self.f_sketches.winfo_children():
            widget.destroy()
        self.f_planes.config(height = 40)
        self.f_sketches.config(height = 40)

        # Get current document, if it is a Part document
        if len(list(self.app.Documents)) == 0:
            return
        doc = self.app.ActiveDocument
        if doc.Type != self.constants.igPartDocument:
            return
        self.doc = doc

        # Activate/deactivate active sketch
        if self.app.ActiveEnvironment == "LayoutInPart":
            self.rb_current.config(state = "normal")
        else:
            self.rb_current.config(state = "disabled")
            if self.v_placement.get() == "current":
                self.v_placement.set("new")

        # We have a part document. Load planes and sketches
        self.load_planes()
        self.load_sketches()
        self.on_placement_change()

    def load_planes(self) -> None:
        """Load planes from the active document"""
        current_plane = self.v_planes.get()
        for i, plane in enumerate(self.doc.RefPlanes):
            name = named_planes.get(i, plane.Name)  # For certain plane get pre-determined name (base planes)
            radio_button = ttk.Radiobutton(self.f_planes, text = name, variable = self.v_planes, value = name)
            radio_button.pack(padx = 4, anchor = "w")
            self.planes[name] = plane
            if i == 0:
                self.v_planes.set(name)

        if current_plane in self.planes:
            self.v_planes.set(current_plane)

    def load_sketches(self) -> None:
        """Load sketches from the active document"""
        current_sketch = self.v_sketches.get()
        for i, sketch in enumerate(self.doc.Sketches):
            name = sketch.Name
            radio_button = ttk.Radiobutton(self.f_sketches, text = name, variable = self.v_sketches, value = name)
            radio_button.pack(padx = 4, anchor = "w")
            self.sketches[name] = sketch
            if i == 0:
                self.v_sketches.set(name)

        if current_sketch in self.sketches:
            self.v_sketches.set(current_sketch)

    def load_aerofoil(self) -> None:
        """Load aerofoil data"""
        file = filedialog.askopenfilename(filetypes = [("Aerofoil .dat", ".dat")])
        if not file:
            return
        self.l_selected.config(text = os.path.split(file)[1])

        # I don't know much about aerofoils, but when just looking around I have found some aerofoil .DAT files
        # which have the top and bottom curves split by an extra new line. Most of the ones I have seen are
        # one continuous curve. To deal with the split-in-two first check that there are exactly two segments,
        # then compare that the starting point of the first one is the same as the starting point of the other segment.
        # If this check passes, reverse the first segment, remove the common point and join it with the second segment.
        with open(file, encoding = "utf-8") as f:
            data = f.read()

        segments = []
        for segment in data.split("\n\n"):
            coordinates = aerofoil_match.findall(segment)
            if not coordinates:
                continue
            segments.append([(float(x), float(y)) for x, y in coordinates])

        if len(segments) == 1:
            self.aerofoil = segments[0]

        elif len(segments) == 2 and segments[0][0] == segments[1][0]:
            self.aerofoil = list(reversed(segments[0][1:])) + segments[1]

        else:
            self.aerofoil = []
            messagebox.showwarning("Unknown DAT file", "Unknown format of .DAT file.\nCannot import aerofoil.")

    def get_se_sketch_profile(self):
        """Get/create sketch to draw the aerofoil in"""
        # Active sketch
        if self.v_placement.get() == "current":
            return self.doc.ActiveSketch

        # Draw in new sketch
        if self.v_placement.get() == "new":
            return self.doc.Sketches.AddByPlane(self.planes[self.v_planes.get()]).Profile

        # Draw in selected sketch
        return self.sketches[self.v_sketches.get()].Profile

    def get_transformed_aerofoil(self) -> list:
        """Scale, mirror, rotate, move aerofoil based on user inputs"""
        aerofoil = self.aerofoil

        # Scale
        if self.v_scale.get() in ("width", "height"):
            if self.v_scale.get() == "width":
                axis = 0
                target = self.e_scale_x.get()
            else:
                axis = 1
                target = self.e_scale_y.get()
            min_val = min(aerofoil, key = lambda x: x[axis])[axis]
            max_val = max(aerofoil, key = lambda x: x[axis])[axis]
            scale = target / (1000 * (max_val - min_val))
            aerofoil = [(x * scale, y * scale) for x, y in aerofoil]

        # Mirror
        if self.v_mirror.get() in ("horizontal", "vertical",):
            m_x, m_y = (-1, 1) if self.v_mirror.get() == "horizontal" else (1, -1)
            aerofoil = [(x * m_x, y * m_y) for x, y in aerofoil]

        # Rotate
        angle = self.e_rotate.get()
        if angle:
            sin = math.sin(math.radians(angle))
            cos = math.cos(math.radians(angle))
            aerofoil = [(x * cos + y * sin, -x * sin + y * cos) for x, y in aerofoil]

        # Move
        move_x = self.e_move_x.get() / 1000
        move_y = self.e_move_y.get() / 1000
        if move_x or move_y:
            aerofoil = [(x + move_x, y + move_y) for x, y in aerofoil]

        return [component for point in aerofoil for component in point]

    def import_into_se(self, *_) -> None:
        """Get transformed aerofoil and import it into solid edge"""
        if self.doc is None:
            messagebox.showwarning("No part document", "Can import aerofoils only into part documents.")
            return

        if not self.aerofoil:
            messagebox.showwarning("No aerofoil", "No aerofoil is loaded.")
            return

        aerofoil = self.get_transformed_aerofoil()
        profile = self.get_se_sketch_profile()
        splines = profile.BSplineCurves2d
        splines.AddByPoints(4, len(self.aerofoil), aerofoil)

        # Reload GUI window
        self.reload_se()


def main():

    # Get app and constants
    try:
        app = win32com.client.GetActiveObject("SolidEdge.Application")
        app = COMWrapper(app)
    except com_error:
        messagebox.showwarning("Solid Edge", "Solid Edge must be running.")
        return

    constants = win32com.client.gencache.EnsureModule("{C467A6F5-27ED-11D2-BE30-080036B4D502}", 0, 1, 0).constants
    constants = COMWrapper(constants)

    root = tk.Tk()
    root.title("Solid Edge aerofoil importer")
    root.resizable(False, False)

    MainApplication(root, app, constants).pack(expand = True, fill = "both")
    root.mainloop()


if __name__ == "__main__":
    main()

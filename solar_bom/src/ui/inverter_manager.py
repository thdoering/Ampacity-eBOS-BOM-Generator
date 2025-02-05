import tkinter as tk
from tkinter import ttk, filedialog, messagebox
import json
from pathlib import Path
from typing import Optional, Callable, Dict
from ..models.inverter import InverterSpec, MPPTChannel, MPPTConfig

class InverterManager(ttk.Frame):
    def __init__(self, parent, on_inverter_selected: Optional[Callable[[InverterSpec], None]] = None):
        super().__init__(parent)
        self.parent = parent
        self.on_inverter_selected = on_inverter_selected
        self.inverters: Dict[str, InverterSpec] = {}
        
        self.setup_ui()
        self.load_inverters()
        
    def setup_ui(self):
        """Create and arrange UI components"""
        # Main container with padding
        main_container = ttk.Frame(self, padding="10")
        main_container.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))
        
        # Left side - Inverter List
        list_frame = ttk.LabelFrame(main_container, text="Inverter Library", padding="5")
        list_frame.grid(row=0, column=0, padx=5, pady=5, sticky=(tk.W, tk.E, tk.N, tk.S))
        
        self.inverter_listbox = tk.Listbox(list_frame, width=40, height=15)
        self.inverter_listbox.grid(row=0, column=0, padx=5, pady=5)
        self.inverter_listbox.bind('<<ListboxSelect>>', self.on_inverter_select)
        
        button_frame = ttk.Frame(list_frame)
        button_frame.grid(row=1, column=0, padx=5, pady=5)
        
        ttk.Button(button_frame, text="Import OND", command=self.import_ond).grid(row=0, column=0, padx=2)
        ttk.Button(button_frame, text="Delete", command=self.delete_inverter).grid(row=0, column=1, padx=2)
        
        # Right side - Inverter Editor
        editor_frame = ttk.LabelFrame(main_container, text="Inverter Details", padding="5")
        editor_frame.grid(row=0, column=1, padx=5, pady=5, sticky=(tk.W, tk.E, tk.N, tk.S))
        
        # Basic Info
        ttk.Label(editor_frame, text="Manufacturer:").grid(row=0, column=0, padx=5, pady=2, sticky=tk.W)
        self.manufacturer_var = tk.StringVar()
        ttk.Entry(editor_frame, textvariable=self.manufacturer_var).grid(row=0, column=1, padx=5, pady=2, sticky=(tk.W, tk.E))
        
        ttk.Label(editor_frame, text="Model:").grid(row=1, column=0, padx=5, pady=2, sticky=tk.W)
        self.model_var = tk.StringVar()
        ttk.Entry(editor_frame, textvariable=self.model_var).grid(row=1, column=1, padx=5, pady=2, sticky=(tk.W, tk.E))
        
        ttk.Label(editor_frame, text="Rated Power (kW):").grid(row=2, column=0, padx=5, pady=2, sticky=tk.W)
        self.power_var = tk.StringVar()
        ttk.Entry(editor_frame, textvariable=self.power_var).grid(row=2, column=1, padx=5, pady=2, sticky=(tk.W, tk.E))
        
        # MPPT Configuration
        mppt_frame = ttk.LabelFrame(editor_frame, text="MPPT Configuration", padding="5")
        mppt_frame.grid(row=3, column=0, columnspan=2, padx=5, pady=5, sticky=(tk.W, tk.E))
        
        ttk.Label(mppt_frame, text="Configuration:").grid(row=0, column=0, padx=5, pady=2, sticky=tk.W)
        self.mppt_config_var = tk.StringVar(value=MPPTConfig.INDEPENDENT.value)
        config_combo = ttk.Combobox(mppt_frame, textvariable=self.mppt_config_var)
        config_combo['values'] = [t.value for t in MPPTConfig]
        config_combo.grid(row=0, column=1, padx=5, pady=2, sticky=(tk.W, tk.E))
        
        # MPPT Channels
        channels_frame = ttk.LabelFrame(mppt_frame, text="MPPT Channels", padding="5")
        channels_frame.grid(row=1, column=0, columnspan=2, padx=5, pady=5, sticky=(tk.W, tk.E))
        
        # Add channel button
        self.channels_container = ttk.Frame(channels_frame)
        self.channels_container.grid(row=0, column=0, sticky=(tk.W, tk.E))
        
        ttk.Button(channels_frame, text="Add MPPT Channel", 
                  command=self.add_mppt_channel).grid(row=1, column=0, pady=5)
        
        self.channel_frames = []  # Store references to channel frames
        
        # Voltage Limits
        voltage_frame = ttk.LabelFrame(editor_frame, text="Voltage Specifications", padding="5")
        voltage_frame.grid(row=4, column=0, columnspan=2, padx=5, pady=5, sticky=(tk.W, tk.E))
        
        ttk.Label(voltage_frame, text="Max DC Voltage:").grid(row=0, column=0, padx=5, pady=2, sticky=tk.W)
        self.max_dc_voltage_var = tk.StringVar(value="1500")
        ttk.Entry(voltage_frame, textvariable=self.max_dc_voltage_var).grid(row=0, column=1, padx=5, pady=2, sticky=(tk.W, tk.E))
        
        ttk.Label(voltage_frame, text="Startup Voltage:").grid(row=1, column=0, padx=5, pady=2, sticky=tk.W)
        self.startup_voltage_var = tk.StringVar(value="150")
        ttk.Entry(voltage_frame, textvariable=self.startup_voltage_var).grid(row=1, column=1, padx=5, pady=2, sticky=(tk.W, tk.E))
        
        # Save button
        ttk.Button(editor_frame, text="Save Inverter", command=self.save_inverter).grid(row=5, column=0, columnspan=2, pady=10)

    def add_mppt_channel(self):
        """Add a new MPPT channel input frame"""
        channel_idx = len(self.channel_frames)
        frame = ttk.Frame(self.channels_container)
        frame.grid(row=channel_idx, column=0, pady=2, sticky=(tk.W, tk.E))
        
        # Channel inputs
        ttk.Label(frame, text=f"Channel {channel_idx + 1}:").grid(row=0, column=0, padx=2)
        
        current_var = tk.StringVar(value="12")
        ttk.Label(frame, text="Max Current:").grid(row=0, column=1, padx=2)
        ttk.Entry(frame, textvariable=current_var, width=8).grid(row=0, column=2, padx=2)
        
        voltage_min_var = tk.StringVar(value="200")
        ttk.Label(frame, text="Min Voltage:").grid(row=0, column=3, padx=2)
        ttk.Entry(frame, textvariable=voltage_min_var, width=8).grid(row=0, column=4, padx=2)
        
        voltage_max_var = tk.StringVar(value="1000")
        ttk.Label(frame, text="Max Voltage:").grid(row=0, column=5, padx=2)
        ttk.Entry(frame, textvariable=voltage_max_var, width=8).grid(row=0, column=6, padx=2)
        
        power_var = tk.StringVar(value="12000")
        ttk.Label(frame, text="Max Power:").grid(row=0, column=7, padx=2)
        ttk.Entry(frame, textvariable=power_var, width=8).grid(row=0, column=8, padx=2)
        
        inputs_var = tk.StringVar(value="2")
        ttk.Label(frame, text="Inputs:").grid(row=0, column=9, padx=2)
        ttk.Entry(frame, textvariable=inputs_var, width=4).grid(row=0, column=10, padx=2)
        
        # Delete button
        ttk.Button(frame, text="X", width=2,
                  command=lambda: self.delete_mppt_channel(frame)).grid(row=0, column=11, padx=2)
        
        # Store frame and variables
        self.channel_frames.append({
            'frame': frame,
            'vars': {
                'current': current_var,
                'voltage_min': voltage_min_var,
                'voltage_max': voltage_max_var,
                'power': power_var,
                'inputs': inputs_var
            }
        })
        
    def delete_mppt_channel(self, frame):
        """Remove an MPPT channel input frame"""
        # Find and remove frame from list
        for i, ch_frame in enumerate(self.channel_frames):
            if ch_frame['frame'] == frame:
                self.channel_frames.pop(i)
                break
                
        # Destroy frame widget
        frame.destroy()
        
        # Reorder remaining frames
        for i, ch_frame in enumerate(self.channel_frames):
            ch_frame['frame'].grid(row=i, column=0)
            
    def create_inverter_spec(self) -> Optional[InverterSpec]:
        """Create InverterSpec from current UI values"""
        try:
            # Create MPPT channels
            channels = []
            for channel in self.channel_frames:
                vars = channel['vars']
                channels.append(MPPTChannel(
                    max_input_current=float(vars['current'].get()),
                    min_voltage=float(vars['voltage_min'].get()),
                    max_voltage=float(vars['voltage_max'].get()),
                    max_power=float(vars['power'].get()),
                    num_string_inputs=int(vars['inputs'].get())
                ))
            
            # Create inverter spec
            inverter = InverterSpec(
                manufacturer=self.manufacturer_var.get(),
                model=self.model_var.get(),
                rated_power=float(self.power_var.get()),
                max_efficiency=98.0,  # Default value
                mppt_channels=channels,
                mppt_configuration=MPPTConfig(self.mppt_config_var.get()),
                max_dc_voltage=float(self.max_dc_voltage_var.get()),
                startup_voltage=float(self.startup_voltage_var.get()),
                nominal_ac_voltage=400.0,  # Default value
                max_ac_current=40.0,  # Default value
                power_factor=0.99,  # Default value
                dimensions_mm=(1000, 600, 300),  # Default values
                weight_kg=75.0,  # Default value
                ip_rating="IP65"  # Default value
            )
            
            inverter.validate()
            return inverter
            
        except (ValueError, TypeError) as e:
            messagebox.showerror("Error", str(e))
            return None
            
    def save_inverter(self):
        """Save current inverter"""
        inverter = self.create_inverter_spec()
        if inverter:
            name = f"{inverter.manufacturer} {inverter.model}"
            if name in self.inverters:
                if not messagebox.askyesno("Confirm", f"Inverter '{name}' already exists. Overwrite?"):
                    return
                    
            self.inverters[name] = inverter
            self.save_inverters()
            self.update_inverter_list()
            
            if self.on_inverter_selected:
                self.on_inverter_selected(inverter)
                
            messagebox.showinfo("Success", f"Inverter '{name}' saved successfully")
            
    def load_inverters(self):
        """Load saved inverters from JSON file"""
        inverter_path = Path('data/inverters.json')
        if not inverter_path.exists():
            return
            
        try:
            with open(inverter_path, 'r') as f:
                data = json.load(f)
                self.inverters = {
                    name: InverterSpec(
                        **{k: v for k, v in specs.items() if k != 'mppt_channels' and k != 'mppt_configuration'},
                        mppt_channels=[MPPTChannel(**ch) for ch in specs['mppt_channels']],
                        mppt_configuration=MPPTConfig(specs['mppt_configuration'])
                    )
                    for name, specs in data.items()
                }
            self.update_inverter_list()
        except Exception as e:
            messagebox.showerror("Error", f"Failed to load inverters: {str(e)}")
            
    def save_inverters(self):
        """Save inverters to JSON file"""
        inverter_path = Path('data/inverters.json')
        inverter_path.parent.mkdir(exist_ok=True)
        
        data = {
            f"{inverter.manufacturer} {inverter.model}": {
                **inverter.__dict__,
                'mppt_channels': [ch.__dict__ for ch in inverter.mppt_channels],
                'mppt_configuration': inverter.mppt_configuration.value
            }
            for inverter in self.inverters.values()
        }
        
        with open(inverter_path, 'w') as f:
            json.dump(data, f, indent=2)
            
    def update_inverter_list(self):
        """Update the inverter listbox"""
        self.inverter_listbox.delete(0, tk.END)
        for inverter in self.inverters.values():
            self.inverter_listbox.insert(tk.END, f"{inverter.manufacturer} {inverter.model}")
            
    def import_ond(self):
        """Import inverter from OND file"""
        # TODO: Implement OND file parsing
        messagebox.showinfo("Not Implemented", "OND file import not yet implemented")
        
    def delete_inverter(self):
        """Delete selected inverter"""
        selection = self.inverter_listbox.curselection()
        if not selection:
            return
        
        name = self.inverter_listbox.get(selection[0])
        if messagebox.askyesno("Confirm", f"Delete inverter '{name}'?"):
            del self.inverters[name]
            self.save_inverters()
            self.update_inverter_list()
            
    def on_inverter_select(self, event=None):
        """Handle inverter selection"""
        selection = self.inverter_listbox.curselection()
        if not selection:
            return
            
        name = self.inverter_listbox.get(selection[0])
        inverter = self.inverters[name]
        
        # Update UI with selected inverter
        self.manufacturer_var.set(inverter.manufacturer)
        self.model_var.set(inverter.model)
        self.power_var.set(str(inverter.rated_power))
        self.mppt_config_var.set(inverter.mppt_configuration.value)
        self.max_dc_voltage_var.set(str(inverter.max_dc_voltage))
        self.startup_voltage_var.set(str(inverter.startup_voltage))
        
        # Clear existing channels
        for frame in self.channel_frames:
            frame['frame'].destroy()
        self.channel_frames.clear()
        
        # Add channels from inverter
        for channel in inverter.mppt_channels:
            self.add_mppt_channel()
            vars = self.channel_frames[-1]['vars']
            vars['current'].set(str(channel.max_input_current))
            vars['voltage_min'].set(str(channel.min_voltage))
            vars['voltage_max'].set(str(channel.max_voltage))
            vars['power'].set(str(channel.max_power))
            vars['inputs'].set(str(channel.num_string_inputs))
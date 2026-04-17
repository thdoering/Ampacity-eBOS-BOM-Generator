import tkinter as tk
from tkinter import ttk, messagebox


class DCFeederDialog(tk.Toplevel):
    """
    Popup dialog for bulk-editing DC feeder distances and cable sizes
    across all blocks. Supports paste from clipboard (one or two columns).
    """

    DC_FEEDER_SIZES = [
        "2/0 AWG", "3/0 AWG", "4/0 AWG",
        "250 kcmil", "300 kcmil", "350 kcmil", "400 kcmil", "500 kcmil",
        "600 kcmil", "750 kcmil", "1000 kcmil"
    ]

    def __init__(self, parent, blocks: dict, on_save=None):
        """
        parent  : parent window
        blocks  : dict of {block_id: BlockConfig}
        on_save : callback(updated_blocks) called when user clicks Save
        """
        super().__init__(parent)
        self.title("DC Feeder Distances")
        self.resizable(True, True)
        self.grab_set()  # Modal

        self.blocks = blocks
        self.on_save = on_save

        # Sort block IDs alphabetically
        self.sorted_ids = sorted(blocks.keys())

        # Working copies of values so we don't mutate until Save
        self.distance_vars = {}   # block_id -> tk.StringVar
        self.cable_vars = {}      # block_id -> tk.StringVar
        self.parallel_vars = {}   # block_id -> tk.StringVar (parallel sets per pole)

        for block_id in self.sorted_ids:
            block = blocks[block_id]
            dist = getattr(block, 'dc_feeder_distance_ft', 0.0)
            size = getattr(block, 'dc_feeder_cable_size', '4/0 AWG')
            parallel = getattr(block, 'dc_feeder_parallel_count', 1)
            self.distance_vars[block_id] = tk.StringVar(value=str(dist))
            self.cable_vars[block_id] = tk.StringVar(value=size)
            self.parallel_vars[block_id] = tk.StringVar(value=str(parallel))

        self._build_ui()
        self._center_on_parent(parent)

    # ------------------------------------------------------------------
    # UI construction
    # ------------------------------------------------------------------

    def _build_ui(self):
        main = ttk.Frame(self, padding=10)
        main.pack(fill='both', expand=True)

        # Instructions
        instructions = (
            "Edit distances, cable sizes, and parallel sets per pole for each block.\n"
            "Paste from clipboard: copy one to three columns from Excel\n"
            "(col 1 = distance ft, col 2 = cable size, col 3 = parallel sets),\n"
            "select the first row you want to fill, then click Paste."
        )
        ttk.Label(main, text=instructions, justify='left').pack(anchor='w', pady=(0, 8))

        # --- Table frame ---
        table_frame = ttk.Frame(main)
        table_frame.pack(fill='both', expand=True)

        columns = ('block_id', 'distance_ft', 'cable_size', 'parallel_count')
        self.tree = ttk.Treeview(
            table_frame,
            columns=columns,
            show='headings',
            selectmode='browse',
            height=min(len(self.sorted_ids) + 1, 20)
        )
        self.tree.heading('block_id',       text='Block ID')
        self.tree.heading('distance_ft',    text='Distance (ft)')
        self.tree.heading('cable_size',     text='Cable Size')
        self.tree.heading('parallel_count', text='Parallel Sets')

        self.tree.column('block_id',       width=140, anchor='w')
        self.tree.column('distance_ft',    width=110, anchor='center')
        self.tree.column('cable_size',     width=130, anchor='center')
        self.tree.column('parallel_count', width=100, anchor='center')

        vsb = ttk.Scrollbar(table_frame, orient='vertical', command=self.tree.yview)
        self.tree.configure(yscrollcommand=vsb.set)

        self.tree.grid(row=0, column=0, sticky='nsew')
        vsb.grid(row=0, column=1, sticky='ns')
        table_frame.columnconfigure(0, weight=1)
        table_frame.rowconfigure(0, weight=1)

        # Populate rows
        for block_id in self.sorted_ids:
            dist = self.distance_vars[block_id].get()
            size = self.cable_vars[block_id].get()
            parallel = self.parallel_vars[block_id].get()
            self.tree.insert('', 'end', iid=block_id, values=(block_id, dist, size, parallel))

        # Double-click to edit
        self.tree.bind('<Double-1>', self._on_double_click)
        # Select first row by default
        if self.sorted_ids:
            self.tree.selection_set(self.sorted_ids[0])
            self.tree.focus(self.sorted_ids[0])

        # --- Button row ---
        btn_frame = ttk.Frame(main)
        btn_frame.pack(fill='x', pady=(10, 0))

        ttk.Button(btn_frame, text="Paste from Clipboard",
                   command=self._paste_from_clipboard).pack(side='left', padx=(0, 5))
        ttk.Button(btn_frame, text="Clear All Distances",
                   command=self._clear_distances).pack(side='left', padx=(0, 5))

        ttk.Button(btn_frame, text="Cancel",
                   command=self.destroy).pack(side='right', padx=(5, 0))
        ttk.Button(btn_frame, text="Save",
                   command=self._save).pack(side='right')

    # ------------------------------------------------------------------
    # Inline cell editing
    # ------------------------------------------------------------------

    def _on_double_click(self, event):
        """Start inline editing for the clicked cell."""
        region = self.tree.identify_region(event.x, event.y)
        if region != 'cell':
            return

        col = self.tree.identify_column(event.x)   # '#1', '#2', '#3'
        row_id = self.tree.identify_row(event.y)    # block_id string
        if not row_id:
            return

        col_index = int(col.replace('#', '')) - 1   # 0-based
        if col_index == 0:
            return  # Block ID is read-only

        # Get cell bounding box
        bbox = self.tree.bbox(row_id, col)
        if not bbox:
            return
        x, y, w, h = bbox

        if col_index == 1:
            # Distance — plain Entry
            var = self.distance_vars[row_id]
            entry = ttk.Entry(self.tree, textvariable=var, width=12)
            entry.place(x=x, y=y, width=w, height=h)
            entry.focus_set()
            entry.select_range(0, 'end')

            def commit_entry(e=None):
                self._validate_and_store_distance(row_id)
                entry.destroy()

            entry.bind('<Return>',    commit_entry)
            entry.bind('<Tab>',       commit_entry)
            entry.bind('<FocusOut>',  commit_entry)
            entry.bind('<Escape>',    lambda e: entry.destroy())

        elif col_index == 2:
            # Cable size — Combobox
            var = self.cable_vars[row_id]
            combo = ttk.Combobox(
                self.tree, textvariable=var,
                values=self.DC_FEEDER_SIZES,
                state='readonly', width=14
            )
            combo.place(x=x, y=y, width=w, height=h)
            combo.focus_set()

            def commit_combo(e=None):
                self._refresh_row(row_id)
                combo.destroy()

            combo.bind('<<ComboboxSelected>>', commit_combo)
            combo.bind('<FocusOut>',            commit_combo)
            combo.bind('<Escape>',              lambda e: combo.destroy())

        elif col_index == 3:
            # Parallel count — Spinbox
            var = self.parallel_vars[row_id]
            spin = ttk.Spinbox(
                self.tree, textvariable=var,
                from_=1, to=10, increment=1, width=10
            )
            spin.place(x=x, y=y, width=w, height=h)
            spin.focus_set()

            def commit_spin(e=None):
                self._validate_and_store_parallel(row_id)
                spin.destroy()

            spin.bind('<Return>',   commit_spin)
            spin.bind('<Tab>',      commit_spin)
            spin.bind('<FocusOut>', commit_spin)
            spin.bind('<Escape>',   lambda e: spin.destroy())

    def _validate_and_store_distance(self, block_id):
        """Parse the distance var and refresh the row display."""
        raw = self.distance_vars[block_id].get().strip()
        try:
            value = float(raw)
            if value < 0:
                value = 0.0
        except ValueError:
            value = 0.0
        self.distance_vars[block_id].set(str(value))
        self._refresh_row(block_id)

    def _validate_and_store_parallel(self, block_id):
        """Parse the parallel count var and refresh the row display."""
        raw = self.parallel_vars[block_id].get().strip()
        try:
            value = int(raw)
            if value < 1:
                value = 1
            elif value > 10:
                value = 10
        except ValueError:
            value = 1
        self.parallel_vars[block_id].set(str(value))
        self._refresh_row(block_id)

    def _refresh_row(self, block_id):
        """Update the treeview row from the current vars."""
        self.tree.item(block_id, values=(
            block_id,
            self.distance_vars[block_id].get(),
            self.cable_vars[block_id].get(),
            self.parallel_vars[block_id].get()
        ))

    # ------------------------------------------------------------------
    # Paste from clipboard
    # ------------------------------------------------------------------

    def _paste_from_clipboard(self):
        """
        Read clipboard text and fill distances, cable sizes, and parallel counts
        starting from the currently selected row.

        Expected clipboard format (copied from Excel):
          - One row per line
          - Column 1: distance in feet (numeric)
          - Column 2 (optional): cable size string matching DC_FEEDER_SIZES
          - Column 3 (optional): parallel sets per pole (integer)
          - Columns separated by tab
        """
        try:
            raw = self.clipboard_get()
        except tk.TclError:
            messagebox.showwarning("Paste", "Clipboard is empty or unavailable.", parent=self)
            return

        lines = [ln for ln in raw.splitlines() if ln.strip()]
        if not lines:
            messagebox.showwarning("Paste", "No data found in clipboard.", parent=self)
            return

        # Determine start row
        selected = self.tree.selection()
        if selected:
            start_index = self.sorted_ids.index(selected[0])
        else:
            start_index = 0

        applied = 0
        skipped_size = []
        skipped_parallel = []

        for i, line in enumerate(lines):
            row_index = start_index + i
            if row_index >= len(self.sorted_ids):
                break  # Ran out of blocks

            block_id = self.sorted_ids[row_index]
            parts = line.split('\t')

            # Column 1 — distance
            dist_raw = parts[0].strip()
            try:
                dist_val = float(dist_raw)
                if dist_val < 0:
                    dist_val = 0.0
                self.distance_vars[block_id].set(str(dist_val))
            except ValueError:
                pass  # Leave existing value if can't parse

            # Column 2 — cable size (optional)
            if len(parts) >= 2:
                size_raw = parts[1].strip()
                if size_raw in self.DC_FEEDER_SIZES:
                    self.cable_vars[block_id].set(size_raw)
                elif size_raw:
                    skipped_size.append(f"Row {i+1}: '{size_raw}'")

            # Column 3 — parallel count (optional)
            if len(parts) >= 3:
                parallel_raw = parts[2].strip()
                if parallel_raw:
                    try:
                        pval = int(float(parallel_raw))  # accept "2" or "2.0"
                        if pval < 1:
                            pval = 1
                        elif pval > 10:
                            pval = 10
                        self.parallel_vars[block_id].set(str(pval))
                    except ValueError:
                        skipped_parallel.append(f"Row {i+1}: '{parallel_raw}'")

            self._refresh_row(block_id)
            applied += 1

        msg = f"Pasted {applied} row(s)."
        if skipped_size:
            msg += f"\n\nUnrecognized cable sizes (left unchanged):\n" + "\n".join(skipped_size)
        if skipped_parallel:
            msg += f"\n\nUnrecognized parallel counts (left unchanged):\n" + "\n".join(skipped_parallel)
        messagebox.showinfo("Paste Complete", msg, parent=self)

    # ------------------------------------------------------------------
    # Utility buttons
    # ------------------------------------------------------------------

    def _clear_distances(self):
        if not messagebox.askyesno("Clear", "Set all distances to 0?", parent=self):
            return
        for block_id in self.sorted_ids:
            self.distance_vars[block_id].set("0.0")
            self._refresh_row(block_id)

    # ------------------------------------------------------------------
    # Save
    # ------------------------------------------------------------------

    def _save(self):
        """Write working values back to block objects and call on_save."""
        for block_id in self.sorted_ids:
            block = self.blocks[block_id]
            try:
                block.dc_feeder_distance_ft = float(self.distance_vars[block_id].get())
            except ValueError:
                block.dc_feeder_distance_ft = 0.0
            block.dc_feeder_cable_size = self.cable_vars[block_id].get()
            try:
                pval = int(self.parallel_vars[block_id].get())
                if pval < 1:
                    pval = 1
            except (ValueError, TypeError):
                pval = 1
            block.dc_feeder_parallel_count = pval

        if self.on_save:
            self.on_save()
        self.destroy()

    # ------------------------------------------------------------------
    # Helpers
    # ------------------------------------------------------------------

    def _center_on_parent(self, parent):
        self.update_idletasks()
        pw = parent.winfo_rootx()
        py = parent.winfo_rooty()
        pW = parent.winfo_width()
        pH = parent.winfo_height()
        w = self.winfo_width()
        h = self.winfo_height()
        x = pw + (pW - w) // 2
        y = py + (pH - h) // 2
        self.geometry(f"+{x}+{y}")
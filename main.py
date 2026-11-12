import os
import sqlite3
from datetime import date, datetime
from calendar import monthrange

from kivy.app import App
from kivy.metrics import dp
from kivy.uix.boxlayout import BoxLayout
from kivy.uix.gridlayout import GridLayout
from kivy.uix.label import Label
from kivy.uix.textinput import TextInput
from kivy.uix.button import Button
from kivy.uix.spinner import Spinner
from kivy.uix.scrollview import ScrollView
from kivy.uix.popup import Popup

from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment
from openpyxl.utils import get_column_letter

APP_TITLE = "Plant Maintenance Survey"
DB_NAME = "plant_survey.db"
MONTHS = ["Unknown"] + [datetime(2020, m, 1).strftime("%B") for m in range(1, 13)]
YEARS = ["Unknown"] + [str(y) for y in range(2015, 2041)]

SCHEMA = """
CREATE TABLE IF NOT EXISTS plants (
 id INTEGER PRIMARY KEY AUTOINCREMENT, name TEXT NOT NULL UNIQUE
);
CREATE TABLE IF NOT EXISTS floors (
 id INTEGER PRIMARY KEY AUTOINCREMENT, plant_id INTEGER NOT NULL,
 name TEXT NOT NULL, UNIQUE(plant_id, name),
 FOREIGN KEY(plant_id) REFERENCES plants(id)
);
CREATE TABLE IF NOT EXISTS machines (
 id INTEGER PRIMARY KEY AUTOINCREMENT, floor_id INTEGER NOT NULL,
 name TEXT NOT NULL, machine_number TEXT, machine_type TEXT,
 manufacturer TEXT, model TEXT, serial_number TEXT, installation_year TEXT, remarks TEXT,
 FOREIGN KEY(floor_id) REFERENCES floors(id)
);
CREATE TABLE IF NOT EXISTS components (
 id INTEGER PRIMARY KEY AUTOINCREMENT, machine_id INTEGER NOT NULL,
 original_name TEXT NOT NULL, standardized_name TEXT, part_number TEXT,
 specification TEXT, required_qty REAL, installed_qty REAL, stock_qty REAL,
 lifetime_value REAL, lifetime_unit TEXT, install_month TEXT, install_year TEXT,
 condition TEXT, remarks TEXT,
 FOREIGN KEY(machine_id) REFERENCES machines(id)
);
"""

def as_float(value):
    value = (value or "").strip()
    if not value:
        return None
    return float(value)

def estimated_dates(month, year):
    if not month or not year or month == "Unknown" or year == "Unknown":
        return None, None, None
    m = datetime.strptime(month, "%B").month
    y = int(year)
    # Mid-month is used as an estimate, not a confirmed installation date.
    start = date(y, m, 15)
    return start, (date.today() - start).days, start

def show_message(title, message):
    box = BoxLayout(orientation="vertical", padding=dp(12), spacing=dp(10))
    box.add_widget(Label(text=message, halign="left", valign="middle"))
    btn = Button(text="OK", size_hint_y=None, height=dp(44))
    box.add_widget(btn)
    popup = Popup(title=title, content=box, size_hint=(.88, .42))
    btn.bind(on_release=popup.dismiss)
    popup.open()

class SurveyRoot(BoxLayout):
    def __init__(self, **kwargs):
        super().__init__(orientation="vertical", padding=dp(10), spacing=dp(8), **kwargs)
        self.db_path = os.path.join(App.get_running_app().user_data_dir, DB_NAME)
        os.makedirs(os.path.dirname(self.db_path), exist_ok=True)
        self.conn = sqlite3.connect(self.db_path)
        self.conn.execute("PRAGMA foreign_keys = ON")
        self.conn.executescript(SCHEMA)
        self.conn.commit()
        self.selected_plant = None
        self.selected_floor = None
        self.selected_machine = None
        self.build_ui()
        self.refresh_all()

    def field(self, hint, multiline=False):
        return TextInput(hint_text=hint, multiline=multiline, size_hint_y=None,
                         height=dp(44 if not multiline else 76), padding=(dp(10), dp(10)))

    def heading(self, text):
        return Label(text=f"[b]{text}[/b]", markup=True, size_hint_y=None,
                     height=dp(34), halign="left", valign="middle")

    def button(self, text, callback, height=46):
        b = Button(text=text, size_hint_y=None, height=dp(height))
        b.bind(on_release=lambda *_: callback())
        return b

    def build_ui(self):
        top = BoxLayout(size_hint_y=None, height=dp(48))
        top.add_widget(Label(text=f"[b]{APP_TITLE}[/b]", markup=True, font_size="18sp"))
        self.add_widget(top)

        scroll = ScrollView()
        self.body = BoxLayout(orientation="vertical", spacing=dp(8), size_hint_y=None)
        self.body.bind(minimum_height=self.body.setter("height"))
        scroll.add_widget(self.body)
        self.add_widget(scroll)

        self.body.add_widget(self.heading("1. Plant"))
        row = BoxLayout(size_hint_y=None, height=dp(44), spacing=dp(6))
        self.plant_spinner = Spinner(text="Select plant", values=(), size_hint_x=.65)
        self.plant_spinner.bind(text=self.on_plant)
        row.add_widget(self.plant_spinner)
        row.add_widget(self.button("Add Plant", self.add_plant, 44))
        self.body.add_widget(row)

        self.body.add_widget(self.heading("2. Floor"))
        row = BoxLayout(size_hint_y=None, height=dp(44), spacing=dp(6))
        self.floor_spinner = Spinner(text="Select floor", values=(), size_hint_x=.65)
        self.floor_spinner.bind(text=self.on_floor)
        row.add_widget(self.floor_spinner)
        row.add_widget(self.button("Add Floor", self.add_floor, 44))
        self.body.add_widget(row)

        self.body.add_widget(self.heading("3. Machine"))
        self.machine_name = self.field("Machine name (required)")
        self.machine_number = self.field("Machine number / ID (optional)")
        self.machine_type = self.field("Machine type")
        self.machine_manufacturer = self.field("Manufacturer")
        self.machine_model = self.field("Model")
        self.machine_serial = self.field("Serial number")
        self.machine_year = self.field("Installation year (optional)")
        self.machine_remarks = self.field("Machine remarks", True)
        for w in [self.machine_name, self.machine_number, self.machine_type,
                  self.machine_manufacturer, self.machine_model, self.machine_serial,
                  self.machine_year, self.machine_remarks]:
            self.body.add_widget(w)
        self.body.add_widget(self.button("Save Machine", self.save_machine))

        self.body.add_widget(self.heading("4. Component"))
        self.component_name = self.field("Component name as observed (required)")
        self.standard_name = self.field("Standardized component name (optional)")
        self.part_number = self.field("Part number")
        self.specification = self.field("Specification")
        self.required_qty = self.field("Required quantity")
        self.installed_qty = self.field("Installed quantity")
        self.stock_qty = self.field("Available spare stock")
        self.lifetime_value = self.field("Expected lifetime (number)")
        self.lifetime_unit = Spinner(text="Lifetime unit", values=("Days", "Months", "Years"),
                                      size_hint_y=None, height=dp(44))
        self.install_month = Spinner(text="Installation month", values=tuple(MONTHS),
                                      size_hint_y=None, height=dp(44))
        self.install_year = Spinner(text="Installation year", values=tuple(YEARS),
                                     size_hint_y=None, height=dp(44))
        self.condition = Spinner(text="Condition", values=("Unknown", "Good", "Fair", "Poor", "Damaged"),
                                 size_hint_y=None, height=dp(44))
        self.component_remarks = self.field("Component remarks", True)
        for w in [self.component_name, self.standard_name, self.part_number, self.specification,
                  self.required_qty, self.installed_qty, self.stock_qty, self.lifetime_value,
                  self.lifetime_unit, self.install_month, self.install_year, self.condition,
                  self.component_remarks]:
            self.body.add_widget(w)
        self.body.add_widget(self.button("Save Component", self.save_component))

        self.body.add_widget(self.heading("Current survey selection"))
        self.selection_label = Label(text="Select a plant, floor and machine.", size_hint_y=None,
                                     height=dp(58), halign="left", valign="middle")
        self.body.add_widget(self.selection_label)
        self.body.add_widget(self.button("Export all data to Excel", self.export_excel, 52))
        self.body.add_widget(Label(text="Data is saved on this phone. Excel is generated only when you export.",
                                   size_hint_y=None, height=dp(48), color=(.35,.35,.35,1)))

    def refresh_all(self):
        rows = self.conn.execute("SELECT name FROM plants ORDER BY name").fetchall()
        self.plant_spinner.values = [r[0] for r in rows]
        if self.selected_plant:
            self.plant_spinner.text = self.selected_plant
        self.refresh_floors()
        self.update_selection()

    def on_plant(self, spinner, text):
        if text == "Select plant": return
        self.selected_plant = text
        self.selected_floor = None
        self.selected_machine = None
        self.floor_spinner.text = "Select floor"
        self.refresh_floors()
        self.update_selection()

    def refresh_floors(self):
        self.floor_spinner.values = ()
        if not self.selected_plant: return
        row = self.conn.execute("SELECT id FROM plants WHERE name=?", (self.selected_plant,)).fetchone()
        if not row: return
        rows = self.conn.execute("SELECT name FROM floors WHERE plant_id=? ORDER BY name", (row[0],)).fetchall()
        self.floor_spinner.values = [r[0] for r in rows]

    def on_floor(self, spinner, text):
        if text == "Select floor": return
        self.selected_floor = text
        self.selected_machine = None
        self.update_selection()

    def add_plant(self):
        self.prompt_text("Add Plant", "Plant name", self.insert_plant)

    def insert_plant(self, value):
        try:
            self.conn.execute("INSERT INTO plants(name) VALUES(?)", (value.strip(),))
            self.conn.commit()
            self.selected_plant = value.strip()
            self.selected_floor = None
            self.selected_machine = None
            self.refresh_all()
        except sqlite3.IntegrityError:
            show_message("Already exists", "This plant name is already in the database.")

    def add_floor(self):
        if not self.selected_plant:
            show_message("Select plant", "Please select a plant first.")
            return
        self.prompt_text("Add Floor", "Floor name", self.insert_floor)

    def insert_floor(self, value):
        pid = self.conn.execute("SELECT id FROM plants WHERE name=?", (self.selected_plant,)).fetchone()[0]
        try:
            self.conn.execute("INSERT INTO floors(plant_id,name) VALUES(?,?)", (pid, value.strip()))
            self.conn.commit()
            self.selected_floor = value.strip()
            self.refresh_floors()
            self.floor_spinner.text = self.selected_floor
            self.update_selection()
        except sqlite3.IntegrityError:
            show_message("Already exists", "This floor already exists under this plant.")

    def prompt_text(self, title, hint, callback):
        layout = BoxLayout(orientation="vertical", padding=dp(12), spacing=dp(8))
        inp = TextInput(hint_text=hint, multiline=False, size_hint_y=None, height=dp(46))
        layout.add_widget(inp)
        buttons = BoxLayout(size_hint_y=None, height=dp(44), spacing=dp(8))
        popup = Popup(title=title, content=layout, size_hint=(.9, .35))
        save = Button(text="Save")
        cancel = Button(text="Cancel")
        buttons.add_widget(save); buttons.add_widget(cancel)
        layout.add_widget(buttons)
        save.bind(on_release=lambda *_: (callback(inp.text) if inp.text.strip() else None, popup.dismiss()))
        cancel.bind(on_release=popup.dismiss)
        popup.open()

    def save_machine(self):
        if not self.selected_floor:
            show_message("Select floor", "Select a plant and floor before saving a machine.")
            return
        name = self.machine_name.text.strip()
        if not name:
            show_message("Missing information", "Machine name is required.")
            return
        fid = self.conn.execute("""SELECT f.id FROM floors f JOIN plants p ON p.id=f.plant_id
                                   WHERE p.name=? AND f.name=?""",
                                (self.selected_plant, self.selected_floor)).fetchone()[0]
        cur = self.conn.execute("""INSERT INTO machines
            (floor_id,name,machine_number,machine_type,manufacturer,model,serial_number,installation_year,remarks)
            VALUES(?,?,?,?,?,?,?,?,?)""",
            (fid, name, self.machine_number.text.strip(), self.machine_type.text.strip(),
             self.machine_manufacturer.text.strip(), self.machine_model.text.strip(),
             self.machine_serial.text.strip(), self.machine_year.text.strip(), self.machine_remarks.text.strip()))
        self.conn.commit()
        self.selected_machine = cur.lastrowid
        self.clear_machine_fields()
        self.update_selection()
        show_message("Saved", "Machine saved. You can now add its components.")

    def clear_machine_fields(self):
        for w in [self.machine_name, self.machine_number, self.machine_type, self.machine_manufacturer,
                  self.machine_model, self.machine_serial, self.machine_year, self.machine_remarks]:
            w.text = ""

    def save_component(self):
        if not self.selected_machine:
            show_message("Select machine", "Save a machine first, then add its components.")
            return
        name = self.component_name.text.strip()
        if not name:
            show_message("Missing information", "Component name is required.")
            return
        try:
            vals = (self.selected_machine, name, self.standard_name.text.strip(),
                    self.part_number.text.strip(), self.specification.text.strip(),
                    as_float(self.required_qty.text), as_float(self.installed_qty.text),
                    as_float(self.stock_qty.text), as_float(self.lifetime_value.text),
                    "" if self.lifetime_unit.text == "Lifetime unit" else self.lifetime_unit.text,
                    "" if self.install_month.text == "Installation month" else self.install_month.text,
                    "" if self.install_year.text == "Installation year" else self.install_year.text,
                    "" if self.condition.text == "Condition" else self.condition.text,
                    self.component_remarks.text.strip())
            self.conn.execute("""INSERT INTO components
                (machine_id,original_name,standardized_name,part_number,specification,required_qty,
                 installed_qty,stock_qty,lifetime_value,lifetime_unit,install_month,install_year,condition,remarks)
                VALUES(?,?,?,?,?,?,?,?,?,?,?,?,?,?)""", vals)
            self.conn.commit()
            self.clear_component_fields()
            show_message("Saved", "Component saved successfully.")
        except ValueError:
            show_message("Invalid quantity", "Quantity and lifetime fields must contain numbers.")

    def clear_component_fields(self):
        for w in [self.component_name, self.standard_name, self.part_number, self.specification,
                  self.required_qty, self.installed_qty, self.stock_qty, self.lifetime_value,
                  self.component_remarks]:
            w.text = ""
        self.lifetime_unit.text = "Lifetime unit"
        self.install_month.text = "Installation month"
        self.install_year.text = "Installation year"
        self.condition.text = "Condition"

    def update_selection(self):
        self.selection_label.text = (f"Plant: {self.selected_plant or '-'}\\n"
                                      f"Floor: {self.selected_floor or '-'}\\n"
                                      f"Machine: {self.selected_machine or '-'}")

    def export_excel(self):
        wb = Workbook()
        summary = wb.active
        summary.title = "Summary"
        summary.append(["Plant Maintenance & LC Planning Survey"])
        summary.append(["Generated on", date.today().isoformat()])
        summary.append(["Plants", self.conn.execute("SELECT COUNT(*) FROM plants").fetchone()[0]])
        summary.append(["Floors", self.conn.execute("SELECT COUNT(*) FROM floors").fetchone()[0]])
        summary.append(["Machines", self.conn.execute("SELECT COUNT(*) FROM machines").fetchone()[0]])
        summary.append(["Component records", self.conn.execute("SELECT COUNT(*) FROM components").fetchone()[0]])

        headers = ["Plant", "Floor", "Machine Name", "Machine Number", "Machine Type",
                   "Manufacturer", "Model", "Serial Number", "Machine Installation Year",
                   "Original Component Name", "Standardized Component Name", "Part Number",
                   "Specification", "Required Qty", "Installed Qty", "Available Stock",
                   "Expected Lifetime", "Lifetime Unit", "Approx. Installation Month",
                   "Approx. Installation Year", "Estimated Days in Service",
                   "Estimated Remaining Life (days)", "Estimated Replacement Date",
                   "Condition", "Machine Remarks", "Component Remarks"]
        detail = wb.create_sheet("All Components")
        detail.append(headers)
        query = """SELECT p.name,f.name,m.name,m.machine_number,m.machine_type,m.manufacturer,
                   m.model,m.serial_number,m.installation_year,c.original_name,c.standardized_name,
                   c.part_number,c.specification,c.required_qty,c.installed_qty,c.stock_qty,
                   c.lifetime_value,c.lifetime_unit,c.install_month,c.install_year,
                   c.condition,m.remarks,c.remarks
                   FROM components c JOIN machines m ON m.id=c.machine_id
                   JOIN floors f ON f.id=m.floor_id JOIN plants p ON p.id=f.plant_id
                   ORDER BY p.name,f.name,m.name"""
        for row in self.conn.execute(query):
            (plant,floor,mname,mnum,mtype,mfr,model,serial,myear,oname,sname,pn,spec,
             req,inst,stock,life,unit,imonth,iyear,cond,mrem,crem) = row
            days_used = remaining = replacement = None
            start, used, _ = estimated_dates(imonth, iyear)
            if start is not None:
                days_used = max(0, used)
                if life is not None and unit:
                    life_days = life if unit == "Days" else life * (30.4375 if unit == "Months" else 365.25)
                    remaining = round(life_days - days_used, 1)
                    replacement = (start + __import__("datetime").timedelta(days=round(life_days))).isoformat()
            detail.append([plant,floor,mname,mnum,mtype,mfr,model,serial,myear,oname,sname,pn,spec,
                           req,inst,stock,life,unit,imonth,iyear,days_used,remaining,replacement,
                           cond,mrem,crem])

        # Procurement sheet: component names are grouped by standardized name, falling back to original.
        procurement = wb.create_sheet("Procurement Summary")
        procurement.append(["Component", "Part Number", "Total Required Qty", "Total Installed Qty",
                            "Available Stock (entered per machine)", "Indicative Net Requirement"])
        groups = {}
        for row in self.conn.execute("""SELECT COALESCE(NULLIF(c.standardized_name,''),c.original_name),
                 c.part_number,c.required_qty,c.installed_qty,c.stock_qty
                 FROM components c"""):
            key = (row[0], row[1] or "")
            g = groups.setdefault(key, [0,0,0])
            g[0] += row[2] or 0
            g[1] += row[3] or 0
            g[2] += row[4] or 0
        for (name,pn),(req,inst,stock) in sorted(groups.items()):
            procurement.append([name,pn,req,inst,stock,max(0,req-stock)])

        # Machine register and plant/floor register
        machines = wb.create_sheet("Machine Register")
        machines.append(["Plant","Floor","Machine Name","Machine Number","Type","Manufacturer","Model",
                         "Serial Number","Installation Year","Remarks"])
        for row in self.conn.execute("""SELECT p.name,f.name,m.name,m.machine_number,m.machine_type,
                    m.manufacturer,m.model,m.serial_number,m.installation_year,m.remarks
                    FROM machines m JOIN floors f ON f.id=m.floor_id JOIN plants p ON p.id=f.plant_id
                    ORDER BY p.name,f.name,m.name"""):
            machines.append(list(row))

        for ws in wb.worksheets:
            ws.freeze_panes = "A2"
            ws.auto_filter.ref = ws.dimensions if ws.max_row > 1 else "A1"
            for cell in ws[1]:
                cell.font = Font(bold=True, color="FFFFFF")
                cell.fill = PatternFill("solid", fgColor="1F4E78")
                cell.alignment = Alignment(wrap_text=True, vertical="center")
            for col in ws.columns:
                max_len = min(max((len(str(c.value or "")) for c in col), default=10) + 2, 34)
                ws.column_dimensions[get_column_letter(col[0].column)].width = max(12, max_len)
            ws.row_dimensions[1].height = 32

        out_dir = App.get_running_app().user_data_dir
        filename = os.path.join(out_dir, f"Plant_Maintenance_Survey_{date.today().isoformat()}.xlsx")
        wb.save(filename)
        show_message("Excel generated", f"Workbook saved on this device:\\n{filename}\\n\\nUse your phone's file manager to share/copy it.")

class PlantSurveyApp(App):
    title = APP_TITLE
    def build(self):
        return SurveyRoot()

if __name__ == "__main__":
    PlantSurveyApp().run()

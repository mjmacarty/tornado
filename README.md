# Sensitivity Analysis Add-in — Installation Guide
## Tornado Chart & Spider Chart Generator for Excel
### Free to use and distribute

---

## What You Get

| Feature | Details |
|---|---|
| Tornado Chart | Ranked horizontal bar chart showing which inputs drive the most output variation |
| Spider Chart | Line chart showing output sensitivity across a ±40% range for each input |
| Cell selection | Point-and-click to select output and input cells directly in Excel |
| Auto-labeling | Reads adjacent cell labels automatically |
| Up to 20 inputs | Handles most real-world models |
| Cross-platform | Works on Excel for Windows and Mac (Excel 365) |

---

## Option A — Quick Install (Recommended)

1. Download `sensitivity.xlam` 
2. This file should be stored in a trusted, permament directory like Documents
3. Open Excel → **File → Options → Add-ins**
3. At the bottom, set **Manage: Excel Add-ins** → click **Go…**
4. Click **Browse…** → navigate to the `sensitivity.xlam` file → click **OK**
5. Make sure the checkbox next to **Sensitivity Analysis** is ticked → **OK**
6. A new **"Sensitivity"** tab appears in your Excel ribbon

> **Note:** When you first open Excel after installing, you may see a security warning about macros. Click **Enable** to allow the add-in to run. 

> **Note:** You may also need to right-click sensitivity.xlam, select properties, and check the **unblock** check box in properties.

---

## Option B — Build It Yourself from Source

Use this option if you want to inspect, modify, or repackage the code yourself.

### Step 1: Create the .xlam file

1. Open a **new, blank** Excel workbook
2. Press **Alt+F11** (Windows) or **Option+F11** (Mac) to open the VBA editor
3. In the Project Explorer, right-click your project → **Insert → Module**
4. Go to **File → Import File** (Ctrl+M on Windows) and import `Module1.bas`
5. Delete the empty default module created in step 3
6. Close the VBA editor
7. **File → Save As** → choose file type: **Excel Add-in (*.xlam)**
8. Name it `SensitivityAnalysis.xlam` and save

### Step 2: Add the Ribbon button via RibbonX Editor

1. Download the free **Office RibbonX Editor**:
   https://github.com/fernandreu/office-ribbonx-editor/releases
2. **Close Excel completely** — RibbonX Editor requires Excel to be closed to save
3. Open `SensitivityAnalysis.xlam` in RibbonX Editor
4. Click **Insert → Office 2010+ Custom UI Part**
5. Paste the contents of `customUI14.xml` into the editor
6. Click **Validate** — should show no errors
7. Click **Save**
8. Close RibbonX Editor

### Step 3: Install the add-in

Follow the same steps as Option A above.

---

## How to Use

1. Open any Excel model with input cells that drive an output calculation
2. Click the **Sensitivity** tab in the ribbon → **Run Analysis**
3. **Step 1** — Click to select your **output cell** (e.g. NPV, profit, total cost)
4. **Step 2** — Select all **input cells** you want to vary (hold Ctrl to select multiple)
5. **Step 3** — Enter a **default variation %** (e.g. `10` for ±10%)
6. Click **OK** — two new sheets are created in your workbook:
   - **Tornado Chart** — ranked by impact, widest bar = most sensitive input
   - **Spider Chart** — shows output trajectory as each input moves ±40%

---

## Tips

- **Labels**: Put descriptive text in the cell to the left or above each input cell — the tool reads these automatically as chart labels. If none are found it uses the cell address instead.
- **Re-running**: Simply click Run Analysis again — old Tornado/Spider sheets are replaced automatically.
- **Zero-value inputs**: Inputs with a base value of 0 will show no variation (0 × any % = 0). Set these inputs to a small non-zero value before running.
- **Model must recalculate**: The output cell must update via formulas when input cells change. If your model uses manual calculations or VBA to compute results this tool may not work correctly.

---

## How to Read the Charts

### Tornado Chart
- Each bar represents one input variable
- Bar width shows the total output range when that input moves from low to high
- **Wider bar = more sensitive input**
- Red = output when input is at its low value; Blue = output at its high value
- Dashed line = base case output value
- Inputs are ranked with the most sensitive at the top

### Spider Chart
- Each line represents one input variable
- X-axis = % change in the input (-40% to +40%); Y-axis = resulting output value
- **Steeper line = more sensitive input**
- All lines cross at the base case (0% change)

---

## Compatibility

| Platform | Status |
|---|---|
| Excel 365 (Windows) | ✅ Fully supported |
| Excel 365 (Mac) | ✅ Fully supported |
| Excel 2016+ (Windows) | ✅ Should work |
| Excel 2016+ (Mac) | ✅ Should work |
| Google Sheets | ❌ Not supported |

---

## License

This tool is free to use, modify, and distribute. No attribution required.
If you improve it, consider sharing your changes!

---

## Troubleshooting

**"Macros are disabled" warning**
→ Go to File → Options → Trust Center → Trust Center Settings → Macro Settings → select "Disable all macros with notification", then re-open the file and click Enable when prompted.

**Sensitivity tab doesn't appear after installation**
→ Make sure the add-in is checked in File → Options → Add-ins → Excel Add-ins. Also try closing and fully restarting Excel.

**Charts appear in the wrong workbook**
→ Make sure your model workbook is the active window when you click Run Analysis.

**Output doesn't change when inputs are varied**
→ Check that your output cell is connected to the input cells via formulas. If the output is hardcoded or calculated by a separate macro it won't respond to input changes.

**"Subscript out of range" error**
→ Make sure you selected at least 1 input cell and 1 output cell before proceeding.

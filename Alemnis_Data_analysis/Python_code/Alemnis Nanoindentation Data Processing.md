# Alemnis Nanoindentation Data Processing

## Purpose

This program automatically processes raw Alemnis nanoindentation/compression Excel files (`.xlsx`) and generates publication-quality figures and cleaned datasets.

The program was developed for particle compression experiments performed using the Alemnis nanoindenter.

It performs three levels of processing:

1. **Raw processed data**
2. **Contact-corrected data**
3. **Endpoint-trimmed data**

The figures are generated using

* SciencePlots
* IEEE publication style
* LaTeX rendering

making them directly suitable for journal publications.

---

# Folder Structure

Simply place the Python script together with all Excel files.

Example

```
Experiment Folder/
│
├── process_alemnis_xlsx.py
├── P1.xlsx
├── P2.xlsx
├── P3.xlsx
├── ...
└── PN.xlsx
```

No additional configuration is required.

The program automatically searches for every

```
*.xlsx
```

file in the same folder.

---

# Excel File Structure

Each Excel workbook exported from Alemnis contains one worksheet.

The worksheet contains eight columns.

| Column | Description                 |
| ------ | --------------------------- |
| A      | Corrected displacement time |
| B      | Corrected displacement      |
| C      | Corrected load time         |
| D      | Corrected load              |
| E      | Raw displacement time       |
| F      | Raw displacement            |
| G      | Raw load time               |
| H      | Raw load                    |

The program **uses only the corrected channels (Columns A–D).**

The raw channels are ignored because the corrected signals already include instrument corrections.

---

# Units Inside Excel

The Excel file stores values using engineering prefixes.

Examples

```
655n
```

means

```
655 × 10⁻⁹
```

```
1.42m
```

means

```
1.42 × 10⁻³
```

Possible prefixes

| Suffix | Meaning |
| ------ | ------- |
| f      | 10⁻¹⁵   |
| p      | 10⁻¹²   |
| n      | 10⁻⁹    |
| u (µ)  | 10⁻⁶    |
| m      | 10⁻³    |

The program automatically converts everything into SI units.

Finally it reports

* displacement in **µm**
* load in **mN**
* time in **seconds**

---

# Initial Cleaning

Some Alemnis files contain an initial discontinuity where the time suddenly resets.

Example

```
0.00
0.05
0.10
...
↓

-6.13
-6.08
...
```

This is an instrument artifact.

The program automatically removes the earlier segment and keeps the final continuous measurement.

---

# Time Synchronization

Displacement and load have separate clocks.

The program synchronizes them onto one common time axis.

If necessary,

* load is interpolated onto displacement time.

This produces one consistent dataset

```
Time
↓

Displacement

Load
```

without time mismatch.

---

# Output Folder 1

```
raw_processed_results
```

This contains the processed measurement exactly as exported by the instrument after

* cleaning
* unit conversion
* time synchronization

Nothing else is modified.

Files produced

```
Processed CSV

Displacement vs Time

Load vs Time

Load vs Displacement
```

Both

```
PNG
```

and

```
PDF
```

versions are created.

---

# Contact Correction

During compression,

the indenter approaches the particle.

Initially,

* the indenter moves,

but

* the particle may roll
* the particle may slip
* no real deformation has started.

Therefore the beginning of the recorded displacement is usually **not** the true start of material deformation.

The program automatically detects the first significant load pickup corresponding to actual particle response.

It then

* removes the approach region
* removes slip
* removes rolling motion

and redefines

```
Time = 0

Displacement = 0

Load = 0
```

at the detected contact point.

---

# Contact Detection Algorithm

The algorithm is fully automatic.

Steps

1. Estimate baseline noise.

2. Smooth the signal using a rolling median.

3. Detect statistically significant load increase.

4. Require the increase to persist for several consecutive points.

5. Backtrack to estimate the actual response onset.

The smoothing is used **only for detection**.

The plotted data always remain the measured values.

---

# Output Folder 2

```
contact_corrected_results
```

Contains

```
Processed CSV

Displacement vs Time

Load vs Time

Load vs Displacement

Contact detection diagnostic plot
```

The diagnostic figure shows

* raw load
* smoothed load
* baseline
* detection threshold
* selected contact point

allowing visual verification.

---

# Endpoint Trimming

Some experiments continue recording after

* fracture,
* unloading,
* particle ejection,
* complete loss of contact.

Those data are usually not useful.

Therefore a third processing stage is available.

---

# User Input

After generating the first two folders,

the program asks

```
Approximate terminal displacement (µm)
```

This value does **not** need to be exact.

Example

```
0.18
```

---

# Endpoint Detection

The program then

1. searches only **after peak load**,

2. finds the point nearest the entered displacement,

3. searches nearby for the local load minimum,

4. selects that point as the true experiment end.

The resulting curve contains

```
Contact

↓

Loading

↓

Peak

↓

Failure

↓

End
```

Everything afterwards is discarded.

---

# Output Folder 3

```
endpoint_trimmed_results
```

Contains

```
Processed CSV

Displacement vs Time

Load vs Time

Load vs Displacement

Endpoint detection diagnostic plot
```

The diagnostic figure shows

* search interval,
* approximate displacement,
* selected endpoint.

---

# Generated CSV Files

Each processing stage exports one CSV.

Columns

```
Time (s)

Displacement (µm)

Load (mN)
```

These files are intended for

* MATLAB
* Python
* Origin
* Excel
* publication figures
* machine learning
* statistical analysis

---

# Publication Figures

All plots use

```
SciencePlots

IEEE style

LaTeX rendering
```

Features include

* vector PDF output
* high-resolution PNG
* publication font sizes
* inward ticks
* minor ticks
* grid
* consistent formatting

The figures can normally be inserted directly into journal manuscripts without further modification.

---

# Running the Program

Install required packages

```bash
python -m pip install -r requirements_alemnis.txt
```

Run

```bash
python process_alemnis_xlsx.py
```

or

```bash
python process_alemnis_xlsx.py --end-displacement-um 0.18
```

The second command performs endpoint trimming automatically without prompting for user input.

---

# Summary of Workflow

```
Excel Files
      │
      ▼
Read corrected channels
      │
      ▼
Convert units
      │
      ▼
Remove time discontinuity
      │
      ▼
Synchronize load & displacement
      │
      ▼
───────────────
Raw processed
───────────────
      │
      ▼
Detect true contact
      │
      ▼
Reset zero
      │
      ▼
────────────────────
Contact corrected
────────────────────
      │
      ▼
User supplies approximate endpoint
      │
      ▼
Locate nearby local load minimum
      │
      ▼
────────────────────
Endpoint trimmed
────────────────────
      │
      ▼
CSV + Publication-quality Figures
```

This workflow preserves the original measurement while also producing physically meaningful datasets for mechanical analysis and publication.

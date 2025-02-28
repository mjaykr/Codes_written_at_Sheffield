import pandas as pd
import matplotlib.pyplot as plt
import os
import glob
import datetime

# Define unit conversion dictionary
unit_conversion = {
    'f': 1e-15,  # femto
    'p': 1e-12,  # pico
    'u': 1e-6,   # micro
    'n': 1e-9,   # nano
    'm': 1e-3    # milli
}

# Define column indices (0-based) for processing
unit_columns = [1, 3, 5, 7]  # Columns with units (displacement, load)
time_columns = [0, 2, 4, 6]  # Columns with time data

# Define headers for the DataFrame
headers = [
    'Time_Corrected_Displacement', 'Displacement_Corrected_Displacement',
    'Time_Corrected_Load', 'Load_Corrected_Load',
    'Raw_Time_Raw_Displacement', 'Displacement_Raw_Displacement',
    'Raw_Time_Raw_Load', 'Load_Raw_Load'
]

# Helper function to convert units
def convert_unit(value, unit_map):
    if isinstance(value, str):
        unit = value[-1]
        if unit in unit_map:
            number = float(value[:-1])
            return number * unit_map[unit]
        else:
            return value
    return value

# Helper function to convert time to seconds
def convert_time(value):
    if isinstance(value, str) and ':' in value:
        return pd.to_timedelta(value).total_seconds()
    elif isinstance(value, datetime.time):
        return value.hour * 3600 + value.minute * 60 + value.second + value.microsecond / 1e6
    elif pd.api.types.is_number(value):
        return value * 86400  # Convert days to seconds
    return value

# Function to ensure time is monotonically increasing
def clean_data(df, time_column, filename):
    rows_to_remove = [False] * len(df)
    for i in range(1, len(df)):
        if df[time_column].iloc[i] <= df[time_column].iloc[i-1]:
            rows_to_remove[i-1] = True
    original_row_count = len(df)
    df = df[~pd.Series(rows_to_remove)].reset_index(drop=True)
    rows_deleted = original_row_count - len(df)
    if rows_deleted > 0:
        print(f'In file {filename}, {rows_deleted} rows deleted to ensure {time_column} is monotonically increasing.')
    return df

# Function to zero-correct time based on displacement
def zero_correct_time(df, time_column, displacement_column):
    displacement = df[displacement_column]
    time = df[time_column]
    
    if (displacement >= 0).any():
        first_non_neg_idx = (displacement >= 0).idxmax()
        time_shift = time.loc[first_non_neg_idx]
    else:
        last_zero_idx = displacement[displacement == 0].index[-1] if (displacement == 0).any() else None
        if last_zero_idx is not None:
            time_shift = time.loc[last_zero_idx]
        else:
            time_shift = time.iloc[0]
    df[time_column] = time - time_shift
    return df

# Function to convert units
def convert_units(df, displacement_column, load_column):
    df[displacement_column] = df[displacement_column] * 1e6  # meters to micrometers
    df[load_column] = df[load_column] * 1e3  # Newtons to milli-Newtons
    return df

# Function to plot and save combined subplots
def plot_combined_subplots(df, filename):
    base_name = os.path.splitext(filename)[0]
    
    plt.rcParams['font.family'] = 'serif'
    plt.rcParams['font.size'] = 14
    plt.rcParams['axes.linewidth'] = 1.5
    plt.rcParams['xtick.direction'] = 'in'
    plt.rcParams['ytick.direction'] = 'in'
    plt.rcParams['xtick.top'] = True
    plt.rcParams['ytick.right'] = True
    
    fig, axs = plt.subplots(1, 3, figsize=(18, 6))
    
    axs[0].plot(df['Time_Corrected_Displacement'], df['Load_Corrected_Load'], linewidth=1.5)
    axs[0].set_xlabel('Time, s')
    axs[0].set_ylabel('Load, mN')
    
    axs[1].plot(df['Time_Corrected_Displacement'], df['Displacement_Corrected_Displacement'], linewidth=1.5)
    axs[1].set_xlabel('Time, s')
    axs[1].set_ylabel(r'Displacement, $\mu$m')
    
    axs[2].plot(df['Displacement_Corrected_Displacement'], df['Load_Corrected_Load'], linewidth=1.5)
    axs[2].set_xlabel(r'Displacement, $\mu$m')
    axs[2].set_ylabel('Load, mN')
    
    plt.tight_layout()
    plt.savefig(f'{base_name}_combined_plots.png', dpi=600, bbox_inches='tight')
    plt.close()

# Function to plot and save individual plots
def plot_individual_plots(df, filename):
    base_name = os.path.splitext(filename)[0]
    
    plt.rcParams['font.family'] = 'serif'
    plt.rcParams['font.size'] = 14
    plt.rcParams['axes.linewidth'] = 1.5
    plt.rcParams['xtick.direction'] = 'in'
    plt.rcParams['ytick.direction'] = 'in'
    plt.rcParams['xtick.top'] = True
    plt.rcParams['ytick.right'] = True
    
    plt.figure(figsize=(6, 4))
    plt.plot(df['Time_Corrected_Displacement'], df['Load_Corrected_Load'], linewidth=1.5)
    plt.xlabel('Time, s')
    plt.ylabel('Load, mN')
    plt.savefig(f'{base_name}_time_vs_load.png', dpi=600, bbox_inches='tight')
    plt.close()
    
    plt.figure(figsize=(6, 4))
    plt.plot(df['Time_Corrected_Displacement'], df['Displacement_Corrected_Displacement'], linewidth=1.5)
    plt.xlabel('Time, s')
    plt.ylabel(r'Displacement, $\mu$m')
    plt.savefig(f'{base_name}_time_vs_displacement.png', dpi=600, bbox_inches='tight')
    plt.close()
    
    plt.figure(figsize=(6, 4))
    plt.plot(df['Displacement_Corrected_Displacement'], df['Load_Corrected_Load'], linewidth=1.5)
    plt.xlabel(r'Displacement, $\mu$m')
    plt.ylabel('Load, mN')
    plt.savefig(f'{base_name}_displacement_vs_load.png', dpi=600, bbox_inches='tight')
    plt.close()

# Function to save displacement and load data
def save_displacement_load_data(df, filename):
    base_name = os.path.splitext(filename)[0]
    data_to_save = df[['Displacement_Corrected_Displacement', 'Load_Corrected_Load']]
    data_to_save.to_csv(f'{base_name}_displacement_load.txt', sep='\t', index=False)

# Function to plot adjusted displacement vs load
def plot_adjusted_displacement_load(df, filename):
    base_name = os.path.splitext(filename)[0]
    
    # Calculate adjusted values by subtracting the first value
    adjusted_displacement = df['Displacement_Corrected_Displacement'] - df['Displacement_Corrected_Displacement'].iloc[0]
    adjusted_load = df['Load_Corrected_Load'] - df['Load_Corrected_Load'].iloc[0]
    
    plt.rcParams['font.family'] = 'serif'
    plt.rcParams['font.size'] = 14
    plt.rcParams['axes.linewidth'] = 1.5
    plt.rcParams['xtick.direction'] = 'in'
    plt.rcParams['ytick.direction'] = 'in'
    plt.rcParams['xtick.top'] = True
    plt.rcParams['ytick.right'] = True
    
    plt.figure(figsize=(6, 4))
    plt.plot(adjusted_displacement, adjusted_load, linewidth=1.5)
    plt.xlabel(r'Displacement, $\mu$m')
    plt.ylabel('Load, mN')
    plt.savefig(f'{base_name}_adjusted_displacement_load.png', dpi=600, bbox_inches='tight')
    plt.close()

# Main execution
excel_files = glob.glob('*.xlsx')

for filename in excel_files:
    df = pd.read_excel(filename, header=None)
    
    for col in unit_columns:
        df[col] = df[col].apply(lambda x: convert_unit(x, unit_conversion))
    
    for col in time_columns:
        df[col] = df[col].apply(convert_time)
    
    df = df.iloc[1:].reset_index(drop=True)
    df.columns = headers
    
    df['Time_Corrected_Displacement'] = pd.to_numeric(df['Time_Corrected_Displacement'], errors='coerce')
    df = df.dropna(subset=['Time_Corrected_Displacement'])
    
    df = clean_data(df, 'Time_Corrected_Displacement', filename)
    df = zero_correct_time(df, 'Time_Corrected_Displacement', 'Displacement_Corrected_Displacement')
    df = convert_units(df, 'Displacement_Corrected_Displacement', 'Load_Corrected_Load')
    
    plot_combined_subplots(df, filename)
    plot_individual_plots(df, filename)
    
    # New additions
    save_displacement_load_data(df, filename)
    plot_adjusted_displacement_load(df, filename)

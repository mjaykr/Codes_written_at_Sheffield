# Automated Analysis and Visualization of Alemnis Experimental Data in MATLAB

### Author: [Mirtunjay Kumar](https://www.linkedin.com/in/mjaykr/)

Understanding the intricacies of material behavior under various loading conditions is crucial for advancements in materials science and engineering. Alemnis, a company specializing in high-precision nanoindentation instruments, provides detailed experimental data that offers insights into material properties at microscopic scales. The data from such instruments are typically comprehensive, necessitating robust analysis methods to derive meaningful conclusions. 

**Detailed Overview of Data Structure in Alemnis Experiments**

Alemnis Standard Assembly (ASA) experiments typically generate comprehensive datasets that offer a snapshot into the mechanical properties of materials at a microscale level. The data structure of such experiments is captured in Excel spreadsheets, meticulously organized into several columns, each representing a different aspect of the experiment. Here's a breakdown of what each column usually signifies:

1. **Time Columns (1, 3, 5, and 7)**: These columns record the time at which the corresponding measurements (like displacement or load) are taken. The time data could be formatted in several ways:
    - **Negative Time Values**: Indicative of pre-test calibrations or specific test protocols that might require the measurement to commence from a negative timestamp, leading up to zero to mark the beginning of the actual test.
    - **Positive Time Values**: These usually follow the negative values and represent the ongoing measurements during the test. The actual data present there is in days formatted to show in hh:mm:ss type. Second unit may extend upto 3 decimal place.

2. **Displacement and Load Measurement Columns (2, 4, 6, and 8)**: These columns carry the crux of the experimental data:
    - **Displacement Data (Columns 2 and 6)**: They may include a prefix denoting the scale (pico, micro, nano, and milli) to represent the magnitude of the displacement measured during the indentation process.
    - **Load Data (Columns 4 and 8)**: Similar to displacement, the load is also recorded with appropriate unit prefixes. This data is crucial as it indicates the force applied by the indenter onto the material surface.

The Alemnis data Excel sheet typically has the following structure:

- **Column 1 (Time_Corrected_Displacement)**: Time data (in second) related to the corrected displacement measurement.
- **Column 2 (Displacement_Corrected_Displacement)**: The corrected displacement values (in meter) that have been measured during the experiment.
- **Column 3 (Time_Corrected_Load)**: Time data (in second) corresponding to the corrected load measurements.
- **Column 4 (Load_Corrected_Load)**: The corrected load values (in Newton) that reflect the force applied to the material.
- **Column 5 (Raw_Time_Raw_Displacement)**: Raw time data (in second) associated with the raw displacement measurements.
- **Column 6 (Displacement_Raw_Displacement)**: Raw displacement data (in meter) as directly measured by the instrument.
- **Column 7 (Raw_Time_Raw_Load)**: Raw time data (in second) related to the raw load measurements.
- **Column 8 (Load_Raw_Load)**: The raw load data (in Newton) indicating the force readings from the device.

The dataset often requires preprocessing to handle the varied time formats and to standardize the units for displacement and load. Negative times are adjusted so that zero corresponds to the beginning of the material deformation, and unit prefixes are standardized to a single unit for consistent analysis. 

In scientific experiments, such as those performed with Alemnis instruments, accuracy in data recording and interpretation is paramount. The Excel file structure provided by Alemnis, with its detailed columns, allows for a granular view of the material's behavior under test conditions. Each column's data is not just a series of numbers; it is a narrative of how a material reacts when subjected to specific conditions, making it invaluable for material characterization and research.

**Adopted Algorithm for Data Analysis**

The algorithm adopted for data analysis of Alemnis experimental results encompasses several steps, from initial data import to final visualization and storage. This structured approach ensures that the data is not only correctly interpreted but also presented in a form that's suitable for further scientific analysis. Here’s a detailed breakdown of the algorithm:

1. **Data Importing**:
   - **Excel File Reading**: The algorithm begins by identifying and reading the Excel files containing the experimental data. Using MATLAB's `xlsread` function, the data from each file is imported into MATLAB's workspace. The function is carefully utilized to avoid the automatic conversion of negative time values to `NaN`.

2. **Data Preprocessing**:
   - **Preallocation**: A new variable mirroring the size of the raw data is preallocated to store processed results, ensuring that the original data remains unaltered.
   - **Unit Conversion**: For columns with displacement and load data, the algorithm identifies the unit prefixes and converts these values into a consistent unit, typically meters for displacement and Newtons for load. This step involves parsing the strings, identifying the unit suffix, and applying the corresponding conversion factor.

3. **Time Correction**:
   - **Normalization of Time Data**: The algorithm checks the `Time_Corrected_Displacement` column for the pattern of zero or positive values followed by negative values. It then eliminates the rows with initial positive values to standardize the starting point of the experiment.
   - **Zero Correction**: If the `Displacement_Corrected_Displacement` column begins with negative values, the algorithm applies a zero correction to the `Time_Corrected_Displacement` column such that the first measurement starts at zero. If multiple zeros are found, the correction is applied from the last zero value before a non-zero entry.

4. **Unit Standardization**:
   - **Displacement Conversion**: All displacement values are converted from meters to micrometers by multiplying by \(10^6\), which is suitable for the scale of the measurements typically taken by Alemnis instruments.
   - **Load Conversion**: The load values are converted from Newtons to milli-Newtons, also by multiplying by a factor of \(10^3\), which aligns with the common units used in material testing.

5. **Iterative Processing**:
   - **Batch File Processing**: The algorithm is designed to loop through all `.xlsx` files in the current directory, applying the entire processing sequence to each file. This batch processing capability is crucial for efficiency when dealing with multiple datasets.

6. **Data Visualization**:
   - **Plotting**: The core of the visualization involves creating plots for the processed data. The algorithm generates plots to illustrate the relationship between displacement and load, as well as the temporal progression of these measurements.
   - **Plot Beautification**: Each plot is enhanced with appropriate labels, including LaTeX-formatted symbols for scientific accuracy, and styled to meet the standards of scientific publication, such as adjusting label positions to prevent overlap with titles and setting figure properties for better clarity.

7. **Output Handling**:
   - **Saving Results**: After processing, the algorithm saves the modified data back into new Excel files, appending a suffix to distinguish them from the raw data files. It also ensures that the figures generated are saved in a high-resolution format suitable for inclusion in scientific documents.

By following this algorithm, researchers can systematically process, analyze, and visualize the data from Alemnis experiments, resulting in a workflow that is both robust and repeatable, essential for the rigorous demands of materials science research.

**Visualization in MATLAB**

The graphical representation of data is pivotal for interpretation and communication in scientific endeavors. In this program, three plots are generated to visualize the relationship between displacement, load, and time:

1. **Displacement vs. Load Plot**: Illustrates how the load carried by a material changes with displacement, giving insights into material stiffness and yielding behavior.

2. **Time vs. Displacement Plot**: Displays the displacement over time, which can be indicative of creep or relaxation properties of the material.

3. **Time vs. Load Plot**: Shows how the load varies over time, which is critical for understanding the dynamic loading response of the material.

Each plot is meticulously formatted for clarity and aesthetic appeal, adhering to standards suitable for scientific publication. Labels are carefully placed to avoid overlap with data points or titles, and the use of LaTeX ensures that symbols and units are presented accurately.

**Summary**

The MATLAB program developed for the analysis of Alemnis experimental data provides a streamlined and automated approach to handle complex datasets. By combining data preprocessing, unit conversions, and sophisticated plotting, the program facilitates a deeper understanding of material behavior. Such tools are indispensable for researchers and engineers as they translate raw experimental data into actionable insights, pushing the boundaries of material science research.

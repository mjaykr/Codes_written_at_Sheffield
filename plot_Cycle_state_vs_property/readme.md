## State of Cycle vs Property Scatter Plot Generator

This MATLAB script generates publication-quality scatter plots for analyzing material properties across different states of cycle. It's designed to work with experimental data stored in Excel files, particularly for studies involving particle orientations and sizes.

 Features

- Data Import: Reads data from user-selected Excel files.
- Flexible Property Selection: Allows users to choose which property to plot against the state of cycle.
- Orientation-based Visualization: Uses different markers and colors to distinguish between particle orientations (e.g., parallel or perpendicular to the basal plane).
- Particle Size Annotation: Option to annotate each data point with its corresponding particle size.
- Smart Label Placement: Implements an algorithm to reduce overlapping of particle size annotations.
- Publication-Ready Formatting: Generates plots with high-quality formatting suitable for academic publications.

 How to Use

1. Ensure your Excel file has a header row with columns for 'State of cycle', 'Orientation', 'Particle Size', and other properties.
2. Run the script in MATLAB.
3. Select your Excel data file when prompted.
4. Choose the property you want to plot from the list of available properties.
5. Decide whether to annotate particle sizes on the plot.

 Key Functions

- Data Reading and Processing: Imports data from Excel and structures it for plotting.
- Interactive User Input: Prompts for file selection and plot customization.
- Scatter Plot Generation: Creates a scatter plot with customized markers and colors.
- Text Annotation: Optionally adds particle size labels to each data point.
- adjust_text(): A helper function that intelligently adjusts text positions to minimize overlap.

 Customization

The script includes several parameters that can be adjusted for different datasets:

- Marker types and colors
- Font sizes and styles
- Axis labels and title
- Figure dimensions
- Text placement algorithm parameters

 Requirements

- MATLAB (developed and tested on version R2021a)
- Excel file with appropriate data structure

 Contributing

Contributions to improve the script or extend its functionality are welcome. Please feel free to fork the repository and submit pull requests.

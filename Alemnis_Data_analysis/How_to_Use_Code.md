# **Short Notes on Using the MATLAB Code for Alemnis Data Analysis**
### Author: [Mirtunjay Kumar](https://www.linkedin.com/in/mjaykr/)

1. **Initial Setup**:
   - Ensure that MATLAB is installed on your system.
   - Place the `.xlsx` files from the Alemnis experiment in a directory that MATLAB can access.

2. **Running the Code**:
   - Open MATLAB and navigate to the directory containing your `.xlsx` files and Matlab Script.
   - Run the MATLAB script. The script will automatically detect and process each `.xlsx` file in the directory.

3. **Data Import and Preprocessing**:
   - The code begins by importing data from the Excel files.
   - It then preallocates a new variable to store processed data, ensuring original data remains unchanged.

4. **Unit Conversion and Time Correction**:
   - Displacement and load data are converted to standard units (micrometers and milli-Newtons).
   - Time data is normalized, with adjustments made for negative or zero starting times.

5. **Visualization and Plotting**:
   - The script generates plots for key relationships (like displacement vs. load).
   - Adjustments are made for clear visualization: setting axis labels (with LaTeX for scientific notation), adjusting line properties, and ensuring no overlap of titles and labels.

6. **Looping Through Files**:
   - The script processes all `.xlsx` files in the folder sequentially.
   - Processed data from each file is saved as a new Excel file with a modified name for distinction.

7. **Output and Saving Figures**:
      - Figures are saved in high-resolution formats suitable for publication.

8. **Post-Processing**:
   - After running the script, you may need to close open figure windows in MATLAB.
   - Review the saved files and figures to ensure they meet your requirements.

9. **Customization**:
   - You can modify the script for specific needs, like changing plot styles or adding additional analysis steps.

10. **Troubleshooting**:
    - If you encounter errors, check the Excel file format and content.
    - Ensure that the script is correctly configured for the specific structure of your Excel files.

This procedure provides a streamlined approach to processing and visualizing Alemnis experimental data. The automation aspect of the script makes it efficient for handling multiple files, significantly reducing manual data processing time.

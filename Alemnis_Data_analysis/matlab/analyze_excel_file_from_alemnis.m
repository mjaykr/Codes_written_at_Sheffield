% Main script to process Excel files with experimental data
clear
files = dir('*.xlsx');

% Process each Excel file
for k = 1:length(files)
    filename = files(k).name;
    processFile(filename);
end

function processFile(filename)
    % Process a single Excel file
    [~, name] = fileparts(filename);
    
    % Step 1: Load and preprocess data
    raw_data = loadExcelData(filename);
    
    % Step 2: Convert units and format data
    processed_data = convertUnitsAndFormat(raw_data);
    
    % Step 3: Create data table with headers
    data_table = createDataTable(processed_data);
    
    % Step 4: Clean data by removing inconsistent time rows
    data_table = removeInconsistentTimeRows(data_table, filename);
    
    % Step 5: Apply zero correction
    data_table = applyZeroCorrection(data_table);
    
    % Step 6: Convert from SI to relevant units
    data_table = convertToRelevantUnits(data_table);
    
    % Step 7: Generate plots
    generatePlots(data_table, name);
    
    % Step 8: Save results
    saveResults(data_table, name);
end

function raw_data = loadExcelData(filename)
    % Load data from Excel file
    [~, ~, raw_data] = xlsread(filename);
end

function processed_data = convertUnitsAndFormat(raw_data)
    % Convert units and format data from raw Excel data
    processed_data = raw_data;
    
    % Create unit conversion map
    unit_conversion = containers.Map({'f', 'p', 'u', 'n', 'm'}, [1e-15, 1e-12, 1e-6, 1e-9, 1e-3]);
    
    % Process columns with units (2, 4, 6, 8)
    processed_data = processUnitColumns(processed_data, raw_data, unit_conversion, [2, 4, 6, 8]);
    
    % Process columns with time and numbers (1, 3, 5, 7)
    processed_data = processTimeColumns(processed_data, raw_data, [1, 3, 5, 7]);
    
    return;
end

function processed_data = processUnitColumns(processed_data, raw_data, unit_conversion, columns)
    % Process columns with units
    for col = columns
        for row = 1:size(raw_data, 1)
            if ischar(raw_data{row, col})
                value = raw_data{row, col};
                unit = value(end);
                if isKey(unit_conversion, unit)
                    number = str2double(value(1:end-1));
                    processed_data{row, col} = number * unit_conversion(unit);
                end
            end
        end
    end
end

function processed_data = processTimeColumns(processed_data, raw_data, columns)
    % Process columns with time and numbers
    for col = columns
        for row = 1:size(raw_data, 1)
            if ischar(raw_data{row, col}) && contains(raw_data{row, col}, ':')
                value = raw_data{row, col};
                time = duration(value);
                processed_data{row, col} = seconds(time);
            elseif ~ischar(raw_data{row, col})
                value = raw_data{row, col};
                mins = floor(value * 1440);
                secs = (value * 1440 - mins) * 60;
                processed_data{row, col} = sprintf('%02d:%06.3f', mins, secs);
                time_parts = sscanf(processed_data{row, col}, '%d:%f');
                processed_data{row, col} = time_parts(1) * 60 + time_parts(2);
            end
        end
    end
end

function data_table = createDataTable(processed_data)
    % Create a table from processed data
    headers = {'Time_Corrected_Displacement', 'Displacement_Corrected_Displacement', ...
               'Time_Corrected_Load', 'Load_Corrected_Load', 'Raw_Time_Raw_Displacement', ...
               'Displacement_Raw_Displacement', 'Raw_Time_Raw_Load', 'Load_Raw_Load'};
    data_table = cell2table(processed_data(2:end, :), 'VariableNames', headers);
end

function data_table = removeInconsistentTimeRows(data_table, filename)
    % Remove rows with inconsistent time at the beginning
    originalRowCount = size(data_table, 1);
    rowsToRemove = false(size(data_table, 1), 1);
    
    for i = 2:size(data_table, 1)
        if data_table.Time_Corrected_Displacement(i) <= data_table.Time_Corrected_Displacement(i - 1)
            rowsToRemove(i-1) = true;
        end
    end
    
    data_table(rowsToRemove, :) = [];
    
    rowsDeleted = originalRowCount - size(data_table, 1);
    if rowsDeleted > 0
        fprintf('In file %s, %d rows have been deleted to ensure Time_Corrected_Displacement is monotonically increasing.\n', ...
                filename, rowsDeleted);
    end
end

function data_table = applyZeroCorrection(data_table)
    % Apply zero correction based on displacement values
    initial_displacement_value = data_table.Displacement_Corrected_Displacement(1);
    
    if initial_displacement_value < 0
        correction_value = abs(data_table.Time_Corrected_Displacement(1));
        data_table.Time_Corrected_Displacement = data_table.Time_Corrected_Displacement + correction_value;
        data_table.Time_Corrected_Displacement(1) = 0;
    elseif initial_displacement_value == 0
        last_zero_index = find(data_table.Displacement_Corrected_Displacement, 1, 'first') - 1;
        if isempty(last_zero_index)
            last_zero_index = size(data_table.Displacement_Corrected_Displacement, 1);
        end
        if last_zero_index > 1
            correction_value = abs(data_table.Time_Corrected_Displacement(last_zero_index));
            data_table.Time_Corrected_Displacement = data_table.Time_Corrected_Displacement + correction_value;
            data_table.Time_Corrected_Displacement(last_zero_index) = 0;
        end
    end
end

function data_table = convertToRelevantUnits(data_table)
    % Convert from SI to relevant units
    
    % Convert displacement from meters to micrometers
    meters_to_micrometers = 1e6;
    data_table.Displacement_Corrected_Displacement = ...
        data_table.Displacement_Corrected_Displacement * meters_to_micrometers;
    
    % Convert load from Newtons to milli-Newtons
    newtons_to_milli_newtons = 1e3;
    data_table.Load_Corrected_Load = ...
        data_table.Load_Corrected_Load * newtons_to_milli_newtons;
end

function generatePlots(data_table, name)
    % Generate all required plots
    
    % Create combined plot
    plotCombined(data_table, name);
    
    % Create individual plots
    plotTimeVsDisplacement(data_table);
    plotTimeVsLoad(data_table);
    plotDisplacementVsLoad(data_table);
end

function plotCombined(data_table, name)
    % Create combined plot with three subplots
    figure('Name', name);
    
    % Time vs Load subplot
    subplot(1, 3, 1);
    xData1 = data_table.Time_Corrected_Displacement;
    yData1 = data_table.Load_Corrected_Load;
    plot(xData1, yData1, 'LineWidth', 1.5);
    xLabel1 = 'Time, s';
    yLabel1 = 'Load, mN';
    formatPlot(xLabel1, yLabel1, 'Time vs. Load');
    
    % Time vs Displacement subplot
    subplot(1, 3, 2);
    xData2 = data_table.Time_Corrected_Displacement;
    yData2 = data_table.Displacement_Corrected_Displacement;
    plot(xData2, yData2, 'LineWidth', 1.5);
    xLabel2 = 'Time, s';
    yLabel2 = 'Displacement, µm';
    formatPlot(xLabel2, yLabel2, 'Time vs. Displacement');
    
    % Displacement vs Load subplot
    subplot(1, 3, 3);
    xData3 = data_table.Displacement_Corrected_Displacement;
    yData3 = data_table.Load_Corrected_Load;

    % Filter out negative values
    valid_indices = xData3 >= 0 & yData3 >= 0;
    xData3 = xData3(valid_indices);
    yData3 = yData3(valid_indices);
    
    % Adjust x values to start from 0
    xData3 = xData3 - xData3(1);

    plot(xData3, yData3, 'LineWidth', 1.5);
    xLabel3 = 'Displacement, µm';
    yLabel3 = 'Load, mN';
    formatPlot(xLabel3, yLabel3, 'Displacement vs. Load');
    
    % Adjust figure properties
    set(gcf, 'Color', 'w', 'Position', [100, 100, 1200, 400]);
    
    % Save combined plot
    print('combined_plots', '-dpng', '-r600');
    
    % Save data from each subplot
    % savePlotData(xData1, yData1, xLabel1, yLabel1, [name, '_time_vs_load']);
    % savePlotData(xData2, yData2, xLabel2, yLabel2, [name, '_time_vs_displacement']);
    savePlotData(xData3, yData3, xLabel3, yLabel3, [name, '_displacement_vs_load']);
end

function plotTimeVsDisplacement(data_table)
    % Create Time vs Displacement plot
    figure;
    xData = data_table.Time_Corrected_Displacement;
    yData = data_table.Displacement_Corrected_Displacement;
    
    plot(xData, yData, 'LineWidth', 2);
    xLabel = 'Time, s';
    yLabel = 'Displacement, µm';
    formatPlot(xLabel, yLabel, '');
    print('myplot_time_vs_displacement', '-dpng', '-r600');
    
    % Save the plotted data
    % savePlotData(xData, yData, xLabel, yLabel, 'data_time_vs_displacement');
end

function plotTimeVsLoad(data_table)
    % Create Time vs Load plot
    figure;
    xData = data_table.Time_Corrected_Displacement;
    yData = data_table.Load_Corrected_Load;
    
    plot(xData, yData, 'LineWidth', 2);
    xLabel = 'Time, s';
    yLabel = 'Load, mN';
    formatPlot(xLabel, yLabel, '');
    print('myplot_time_vs_load', '-dpng', '-r600');
    
    % Save the plotted data
    % savePlotData(xData, yData, xLabel, yLabel, 'data_time_vs_load');
end

function plotDisplacementVsLoad(data_table)
    % Create Displacement vs Load plot
    figure;
    xData = data_table.Displacement_Corrected_Displacement;
    yData = data_table.Load_Corrected_Load;
    % Filter out negative values
    valid_indices = xData >= 0 & yData >= 0;
    xData = xData(valid_indices);
    yData = yData(valid_indices);
    
    % Adjust x values to start from 0
    xData = xData - xData(1);
    
    plot(xData, yData, 'LineWidth', 2);
    xLabel = 'Displacement, µm';
    yLabel = 'Load, mN';
    formatPlot(xLabel, yLabel, '');
    print('myplot_load_displacement', '-dpng', '-r600');
    
    % Save the plotted data
    % savePlotData(xData, yData, xLabel, yLabel, 'data_displacement_vs_load');
end

function formatPlot(xLabel, yLabel, titleText)
    % Apply consistent formatting to plots
    xlabel(xLabel, 'Interpreter', 'tex', 'FontSize', 14);
    ylabel(yLabel, 'Interpreter', 'tex', 'FontSize', 14);
    if ~isempty(titleText)
        title(titleText);
    end
    set(gca, 'Box', 'on', 'GridLineStyle', '-', 'LineWidth', 1, ...
        'FontName', 'Times New Roman', 'FontSize', 14, 'TickDir', 'out');
    grid on;
    set(gcf, 'Color', 'w');
end

function saveResults(data_table, name)
    % Save processed data and figures
    new_filename = sprintf('%s_processed.xlsx', name);
    name_of_figure_plot = sprintf('%s_Plots.png', name);
    
    % Update table headers for saving
    customHeaders = {'Time(s)_Corrected_Displacement', 'Displacement(micro_meter)_Corrected_Displacement', ...
                 'Time(s)_Corrected_Load', 'Load(mN)_Corrected_Load', 'Raw_Time(s)_Raw_Displacement', ...
                 'Displacement(micro_meter)_Raw_Displacement', 'Raw_Time(s)_Raw_Load', 'Load(mN)_Raw_Load'};
    data_table.Properties.VariableNames = customHeaders;
    
    % Save table to Excel
    writetable(data_table, new_filename);
    
    % Save figure
    % print(name_of_figure_plot, '-dpng', '-r600');
    
    % Close all figures
    close all;
end
function savePlotData(xData, yData, xLabel, yLabel, filename)
    % Remove units and formatting from labels to make them valid headers
    xHeader = strrep(xLabel, ', ', '_');
    xHeader = strrep(xHeader, ' ', '');
    yHeader = strrep(yLabel, ', ', '_');
    yHeader = strrep(yHeader, ' ', '');
    
    % Create table with the data
    plotData = table(xData, yData, 'VariableNames', {xHeader, yHeader});
    
    % Save to CSV file
    writetable(plotData, [filename, '.csv']);
    
    % Optionally also save as MAT file for easier MATLAB loading later
    % save([filename, '.mat'], 'xData', 'yData', 'xLabel', 'yLabel');
end

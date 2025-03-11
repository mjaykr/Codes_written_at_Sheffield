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
    
    % Step 7: Generate plots with user-defined displacement range
    generatePlotsWithUserRange(data_table, name, filename);
    
    % Step 8: Save results
    saveResults(data_table, name, filename);
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

function generatePlotsWithUserRange(data_table, name, filename)
    % Generate plots with user-defined displacement range
    
    % Get the full data range for preview
    xData = data_table.Displacement_Corrected_Displacement;
    yData = data_table.Load_Corrected_Load;
    
    % Filter out negative values
    valid_indices = xData >= 0 & yData >= 0;
    xData = xData(valid_indices);
    yData = yData(valid_indices);
    
    % Adjust x values to start from 0
    xData = xData - xData(1);
    
    % Get max displacement
    max_displacement = max(xData);
    user_max = max_displacement; % Initial value
    
    % Loop until user is satisfied with the displacement range
    adjust = 'y';
    while strcmpi(adjust, 'y')
        % Create a preview plot for the current range
        fig_preview = figure('Name', [name, ' - Preview'], 'Position', [100, 100, 800, 600]);
        
        % Plot both full data (light) and selected range (dark)
        preview_indices = xData <= user_max;
        plot(xData, yData, 'Color', [0.8, 0.8, 0.8], 'LineWidth', 1);
        hold on;
        plot(xData(preview_indices), yData(preview_indices), 'LineWidth', 1.5);
        hold off;
        
        xlabel('Displacement, µm', 'Interpreter', 'tex', 'FontSize', 14);
        ylabel('Load, mN', 'Interpreter', 'tex', 'FontSize', 14);
        title(sprintf('Preview: Displacement vs. Load (Selected Range: %.2f µm)', user_max), 'FontSize', 16);
        grid on;
        
        % Format the preview plot
        ax = gca;
        set(ax, ...
            'Box', 'on', ...
            'TickDir', 'in', ...
            'TickLength', [0.02 0.02], ...
            'XMinorTick', 'on', ...
            'YMinorTick', 'on', ...
            'LineWidth', 1.5, ...
            'FontName', 'Times New Roman', ...
            'FontSize', 14);
        
        set(gcf, 'Color', 'w');
        
        % Show max displacement information
        text(0.6, 0.9, sprintf('Max displacement: %.2f µm', max_displacement), ...
             'Units', 'normalized', 'FontSize', 12, 'BackgroundColor', [1 1 0.8]);
        text(0.6, 0.85, sprintf('Selected max: %.2f µm', user_max), ...
             'Units', 'normalized', 'FontSize', 12, 'BackgroundColor', [1 1 0.8]);
        
        % Ask if user wants to adjust the displacement range
        adjust = input('Do you want to adjust the displacement range? (y/n): ', 's');
        
        if strcmpi(adjust, 'y')
            % Ask for a new maximum displacement value
            prompt = sprintf('Enter maximum displacement value to plot (µm) [current: %.2f]: ', user_max);
            new_max_str = input(prompt, 's');
            
            if isempty(new_max_str)
                new_max = user_max;
            else
                new_max = str2double(new_max_str);
                % Validate the input
                while isnan(new_max) || new_max <= 0 || new_max > max_displacement
                    if isnan(new_max)
                        fprintf('Invalid input. Please enter a number.\n');
                    elseif new_max <= 0
                        fprintf('Maximum displacement must be positive.\n');
                    else
                        fprintf('Maximum displacement cannot exceed %.2f µm.\n', max_displacement);
                    end
                    new_max_str = input(prompt, 's');
                    if isempty(new_max_str)
                        new_max = user_max;
                        break;
                    else
                        new_max = str2double(new_max_str);
                    end
                end
            end
            
            % Update user_max
            user_max = new_max;
        end
        
        % Close the preview figure
        close(fig_preview);
    end
    
    % User is satisfied with the range, create trimmed data
    trimmed_data = data_table;
    trim_indices = trimmed_data.Displacement_Corrected_Displacement <= user_max;
    trimmed_data = trimmed_data(trim_indices, :);
    
    % Generate full range combined plot first (without trimming)
    createFullRangeCombinedPlot(data_table, name, filename);
    
    % Now generate and save all plots with the final selected range
    createAllPlots(trimmed_data, name, filename, user_max);
    saveAllPlots(filename, user_max);
    saveFinalData(trimmed_data, filename, user_max);
    
    fprintf('All plots generated and saved with maximum displacement of %.2f µm\n', user_max);
end

function createFullRangeCombinedPlot(data_table, name, filename)
    % Create combined plot with three subplots using the full data range
    figure('Name', [name, ' - Full Range Combined Plots'], 'Position', [100, 100, 1200, 400]);
    
    % Time vs Load subplot
    subplot(1, 3, 1);
    xData1 = data_table.Time_Corrected_Displacement;
    yData1 = data_table.Load_Corrected_Load;
    
    % Threshold-based filtering
    threshold = 0.1; % Adjust as needed
    valid_indices = yData1 > threshold;
    xData1 = xData1(valid_indices);
    yData1 = yData1(valid_indices);
    
    % Shift time to start at 0
    if ~isempty(xData1)
        xData1 = xData1 - xData1(1);
    end
    
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
    
    % Add information to the figure title
    sgtitle('Full Range Plots (No Trimming)', 'FontSize', 16, 'FontWeight', 'bold');
    
    % Adjust figure properties
    set(gcf, 'Color', 'w');
end

function createAllPlots(data_table, name, filename, max_displacement)
    % Create all plots with the trimmed data
    
    % Create combined plot with three subplots
    createCombinedPlot(data_table, name, filename, max_displacement);
    
    % Create individual plots
    createTimeVsDisplacement(data_table, filename, max_displacement);
    createTimeVsLoad(data_table, filename, max_displacement);
    createDisplacementVsLoad(data_table, filename, max_displacement);
end

function createCombinedPlot(data_table, name, filename, max_displacement)
    % Create combined plot with three subplots
    figure('Name', [name, ' - Combined Plots'], 'Position', [100, 100, 1200, 400]);
    
    % Time vs Load subplot
    subplot(1, 3, 1);
    xData1 = data_table.Time_Corrected_Displacement;
    yData1 = data_table.Load_Corrected_Load;
    
    % Threshold-based filtering
    threshold = 0.1; % Adjust as needed
    valid_indices = yData1 > threshold;
    xData1 = xData1(valid_indices);
    yData1 = yData1(valid_indices);
    
    % Shift time to start at 0
    if ~isempty(xData1)
        xData1 = xData1 - xData1(1);
    end
    
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
    
    % Add trimming information to the figure title
    sgtitle(sprintf('Plots Trimmed to %.2f µm Displacement', max_displacement), ...
            'FontSize', 16, 'FontWeight', 'bold');
    
    % Adjust figure properties
    set(gcf, 'Color', 'w');
end

function createTimeVsDisplacement(data_table, filename, max_displacement)
    % Create Time vs Displacement plot
    figure('Name', 'Time vs. Displacement', 'Position', [100, 100, 800, 600]);
    xData = data_table.Time_Corrected_Displacement;
    yData = data_table.Displacement_Corrected_Displacement;
    
    plot(xData, yData, 'LineWidth', 2);
    xLabel = 'Time, s';
    yLabel = 'Displacement, µm';
    formatPlot(xLabel, yLabel, '');
    % title(sprintf('Time vs. Displacement (Max: %.2f µm)', max_displacement), 'FontSize', 16);
end

function createTimeVsLoad(data_table, filename, max_displacement)
    % Create Time vs Load plot
    figure('Name', 'Time vs. Load', 'Position', [100, 100, 800, 600]);
    xData = data_table.Time_Corrected_Displacement;
    yData = data_table.Load_Corrected_Load;
    
    % Define a small threshold to ignore insignificant loads
    threshold = 0.1; % Adjust as needed
    valid_indices = yData > threshold;  
    xData = xData(valid_indices);
    yData = yData(valid_indices);
    
    % Shift time to start from 0
    if ~isempty(xData)
        xData = xData - xData(1);
    end
    
    plot(xData, yData, 'LineWidth', 2);
    xLabel = 'Time, s';
    yLabel = 'Load, mN';
    formatPlot(xLabel, yLabel, '');
    % title(sprintf('Time vs. Load (Displacement Max: %.2f µm)', max_displacement), 'FontSize', 16);
end

function createDisplacementVsLoad(data_table, filename, max_displacement)
    % Create Displacement vs Load plot
    figure('Name', 'Displacement vs. Load', 'Position', [100, 100, 800, 600]);
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
    % title(sprintf('Displacement vs. Load (Max: %.2f µm)', max_displacement), 'FontSize', 16);
end

function formatPlot(xLabel, yLabel, titleText)
    xlabel(xLabel, 'Interpreter', 'tex', 'FontSize', 14);
    ylabel(yLabel, 'Interpreter', 'tex', 'FontSize', 14);
    if ~isempty(titleText)
        title(titleText, 'FontSize', 14);
    end
    
    ax = gca;
    set(ax, ...
        'Box', 'on', ...
        'BoxStyle', 'full', ...        % Requires R2019b or newer
        'TickDir', 'in', ...
        'TickLength', [0.02 0.02], ...
        'XMinorTick', 'on', ...
        'YMinorTick', 'on', ...
        'LineWidth', 1.5, ...
        'FontName', 'Times New Roman', ...
        'FontSize', 14, ...
        'XGrid', 'off', ...
        'YGrid', 'off', ...
        'MinorGridLineStyle', 'none');

    set(gcf, 'Color', 'w');
end

function saveAllPlots(filename, max_displacement)
    % Save all open figures
    figHandles = findall(0, 'Type', 'figure');
    
    for i = 1:length(figHandles)
        fig = figHandles(i);
        figure(fig);
        
        % Get the figure name and sanitize it for use in filenames
        figName = get(fig, 'Name');
        figName = regexprep(figName, '[^a-zA-Z0-9_]', '_');
        
        % Create a descriptive filename
        saveFilename = sprintf('%s_%s_max%.2fum.png', filename, figName, max_displacement);
        
        % Save the figure
        print(saveFilename, '-dpng', '-r600');
    end
end

function saveFinalData(data_table, filename, max_displacement)
    % Save the trimmed data to CSV and MAT files
    
    % Save displacement vs load data
    xData = data_table.Displacement_Corrected_Displacement;
    yData = data_table.Load_Corrected_Load;
    
    % Filter out negative values
    valid_indices = xData >= 0 & yData >= 0;
    xData = xData(valid_indices);
    yData = yData(valid_indices);
    
    % Adjust x values to start from 0
    xData = xData - xData(1);
    
    % Create table for CSV export
    exportTable = table(xData, yData, 'VariableNames', {'Displacement_um', 'Load_mN'});
    
    % Create descriptive filenames with max displacement info
    csvFilename = sprintf('%s_displacement_vs_load_max%.2fum.csv', filename, max_displacement);
    excelFilename = sprintf('%s_displacement_vs_load_max%.2fum.xlsx', filename, max_displacement);
    matFilename = sprintf('%s_displacement_vs_load_max%.2fum.mat', filename, max_displacement);
    
    % Save to files
    writetable(exportTable, csvFilename);
    writetable(exportTable, excelFilename);
    save(matFilename, 'xData', 'yData', 'max_displacement');
    
    fprintf('Plots and data saved with maximum displacement of %.2f µm\n', max_displacement);
end

function saveResults(data_table, name, filename)
    % Save processed data
    new_filename = sprintf('%s_processed.xlsx', name);
    
    % Update table headers for saving
    customHeaders = {'Time(s)_Corrected_Displacement', 'Displacement(micro_meter)_Corrected_Displacement', ...
                 'Time(s)_Corrected_Load', 'Load(mN)_Corrected_Load', 'Raw_Time(s)_Raw_Displacement', ...
                 'Displacement(micro_meter)_Raw_Displacement', 'Raw_Time(s)_Raw_Load', 'Load(mN)_Raw_Load'};
    data_table.Properties.VariableNames = customHeaders;
    
    % Save table to Excel
    writetable(data_table, new_filename);
end

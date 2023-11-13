% List all .xlsx files in the current directory
clear
files = dir('*.xlsx');

% Loop over each file
for k = 1:length(files)
             % Read the Excel file. This function call will get the data into a cell array.
            % xlsread is used here without the first two output arguments which are for numeric data and text data separately.
                filename = files(k).name;
                [~, ~, raw_data] = xlsread(filename);

            % Preallocate a new variable for the processed data to avoid modifying the original raw data
            processed_data = raw_data; % This creates a copy of the raw_data to preserve the original
            
            % Prepare a map for unit conversion factors to convert units like pico, micro, nano, milli to their respective multipliers
            unit_conversion = containers.Map({'p', 'u', 'n', 'm'}, [1e-12, 1e-6, 1e-9, 1e-3]);
            
            % Process the data for columns with units (2, 4, 6, 8)
            for col = [2, 4, 6, 8] 
                for row = 1:size(raw_data, 1)
                    % Check if the data is a string (which means it has a unit) and not empty
                    if ischar(raw_data{row, col})
                        value = raw_data{row, col};
                        unit = value(end); % Extract the unit (last character of the string)
                        if isKey(unit_conversion, unit)
                            % Convert the string to a number without the unit and then multiply by the conversion factor
                            number = str2double(value(1:end-1));
                            processed_data{row, col} = number * unit_conversion(unit);
                        end
                    end
                end
            end
            
            % Process the data for columns with time and numbers (1, 3, 5, 7)
            for col = [1, 3, 5, 7] 
                for row = 1:size(raw_data, 1)
                    if ischar(raw_data{row, col}) && contains(raw_data{row, col}, ':')
                        % If the format is a time string, convert it to a duration object and then to seconds
                        value = raw_data{row, col};
                        time = duration(value);
                        processed_data{row, col} = seconds(time);
                    elseif ~ischar(raw_data{row, col})
                        % Convert numeric values to minutes and seconds, then to a time string, and finally to seconds
                        value = raw_data{row, col};
                        mins = floor(value * 1440); % Convert from days to minutes
                        secs = (value * 1440 - mins) * 60; % Convert the remainder to seconds
                        processed_data{row, col} = sprintf('%02d:%06.3f', mins, secs);
                        time_parts = sscanf(processed_data{row, col}, '%d:%f');
                        processed_data{row, col} = time_parts(1) * 60 + time_parts(2);
                    end
                end
            end
            
            % Define the new headers for the table
            headers = {'Time_Corrected_Displacement', 'Displacement_Corrected_Displacement', ...
                       'Time_Corrected_Load', 'Load_Corrected_Load', 'Raw_Time_Raw_Displacement', ...
                       'Displacement_Raw_Displacement', 'Raw_Time_Raw_Load', 'Load_Raw_Load'};
                   
            % Convert the processed data into a table with the specified headers.
            % We exclude the first row as it contains the original headers.
            data_table = cell2table(processed_data(2:end, :), 'VariableNames', headers);
            
            % Assign the table with the converted data to the base workspace with the variable name 'converted_data'.
            % This allows you to access the table from the MATLAB workspace.
            assignin('base', 'converted_data', data_table);
            
            %% Removing the Rows having inconsistent time at the beginning of the Experiments
            % Original number of rows
            originalRowCount = size(data_table, 1);
            
            % Initialize an index to keep track of rows to be removed
            rowsToRemove = false(size(data_table, 1), 1); % Initially, mark all rows as 'false' (not to be removed)
            
            % Loop through the Time_Corrected_Displacement column
            for i = 2:size(data_table, 1)
                % If the current value is less than or equal to the previous one
                if data_table.Time_Corrected_Displacement(i) <= data_table.Time_Corrected_Displacement(i - 1)
                    rowsToRemove(i-1) = true; % Mark this row for removal
                end
            end
            
            % Remove the marked rows
            data_table(rowsToRemove, :) = [];
            
            % Calculate the number of rows deleted
            rowsDeleted = originalRowCount - size(data_table, 1);
            
            % Output the message only if rows have been deleted
            if rowsDeleted > 0
                fprintf('In file %s, %d rows have been deleted to ensure Time_Corrected_Displacement is monotonically increasing.\n', filename, rowsDeleted);
            end
            
            % The modified data_table now has only rows where Time_Corrected_Displacement is monotonically increasing.
            
            %% Zero Correction 
            
            % Get the initial value of 'Displacement_Corrected_Displacement'
            initial_displacement_value = data_table.Displacement_Corrected_Displacement(1);
            
            % Check if 'Displacement_Corrected_Displacement' starts with negative values
            if initial_displacement_value < 0
                % Apply zero correction from the start
                correction_value = abs(data_table.Time_Corrected_Displacement(1));
                data_table.Time_Corrected_Displacement = data_table.Time_Corrected_Displacement + correction_value;
                data_table.Time_Corrected_Displacement(1) = 0; % Set the first cell to zero after correction
            elseif initial_displacement_value == 0
                % Find the last zero value before a non-zero value in 'Displacement_Corrected_Displacement'
                last_zero_index = find(data_table.Displacement_Corrected_Displacement, 1, 'first') - 1;
                if isempty(last_zero_index)
                    % If the entire column is zero, then we take the last index
                    last_zero_index = size(data_table.Displacement_Corrected_Displacement, 1);
                end
                % Check if there are multiple zeros at the beginning
                if last_zero_index > 1
                    % Apply zero correction from the last zero value
                    correction_value = abs(data_table.Time_Corrected_Displacement(last_zero_index));
                    data_table.Time_Corrected_Displacement = data_table.Time_Corrected_Displacement + correction_value;
                    % Set the time of zero displacement to zero
                    data_table.Time_Corrected_Displacement(last_zero_index) = 0;
                end
            end
            % Now the 'Time_Corrected_Displacement' column has been zero-corrected based on 'Displacement_Corrected_Displacement' conditions.
            
            %% Unit Conversions from SI to the relevant unit
            
            % Conversion factor from meters to micrometers
            meters_to_micrometers = 1e6;
            
            % Convert 'Displacement_Corrected_Displacement' from meters to micrometers
            data_table.Displacement_Corrected_Displacement = ...
                data_table.Displacement_Corrected_Displacement * meters_to_micrometers;
            
            % Now the 'Displacement_Corrected_Displacement' column values are in micrometers.
            
            % Conversion factor from Newtons to milli-Newtons
            newtons_to_milli_newtons = 1e3;
            
            % Convert 'Load_Corrected_Load' from Newtons to milli-Newtons
            data_table.Load_Corrected_Load = ...
                data_table.Load_Corrected_Load * newtons_to_milli_newtons;
            
            % Now the 'Load_Corrected_Load' column values are in milli-Newtons.
            
            
            %% Plotting the Data
            %%%%% Creating all figures in one Plot %%%%%%%%%
            % Create a figure window
            [~, name] = fileparts(filename); % Extract the base name without extension
            figure('Name',name);
            % tiledlayout('flow', 'TileSpacing', 'compact', 'Padding', 'compact');
            % First subplot for Displacement vs. Load
            subplot(1, 3, 1); % 1 row, 3 columns, 1st subplot
            plot(data_table.Time_Corrected_Displacement, data_table.Load_Corrected_Load, 'LineWidth', 2);
            xlabel('Time, s', 'Interpreter', 'latex', 'FontSize', 14);
            ylabel('Load, mN', 'Interpreter', 'latex', 'FontSize', 14);
            title('Time vs. Load'); % Optional: Add a title to the subplot
            set(gca, 'Box', 'on', 'GridLineStyle', '-', 'LineWidth', 1, 'FontName', 'Times New Roman', 'FontSize', 14, 'TickDir', 'out');
            grid on;
            
            % Second subplot for Time vs. Displacement
            subplot(1, 3, 2); % 1 row, 3 columns, 2nd subplot
            plot(data_table.Time_Corrected_Displacement, data_table.Displacement_Corrected_Displacement, 'LineWidth', 2);
            xlabel('Time, s', 'Interpreter', 'latex', 'FontSize', 14);
            ylabel('Displacement, $\mu$m', 'Interpreter', 'latex', 'FontSize', 14);
            title('Time vs. Displacement'); % Optional: Add a title to the subplot
            set(gca, 'Box', 'on', 'GridLineStyle', '-', 'LineWidth', 1, 'FontName', 'Times New Roman', 'FontSize', 14, 'TickDir', 'out');
            grid on;
            
            % Third subplot for Time vs. Load
            subplot(1, 3, 3); % 1 row, 3 columns, 3rd subplot
            plot(data_table.Displacement_Corrected_Displacement, data_table.Load_Corrected_Load, 'LineWidth', 2);
            xlabel('Displacement, $\mu$m', 'Interpreter', 'latex', 'FontSize', 14);
            ylabel('Load, mN', 'Interpreter', 'latex', 'FontSize', 14);
            title('Displacement vs. Load'); % Optional: Add a title to the subplot
            set(gca, 'Box', 'on', 'GridLineStyle', '-', 'LineWidth', 1, 'FontName', 'Times New Roman', 'FontSize', 14, 'TickDir', 'out');
            grid on;
            
            % Adjust the figure properties for better spacing
            set(gcf, 'Color', 'w', 'Position', [100, 100, 1200, 400]); % Set the figure size and position
            
            % If you want to save this combined figure for inclusion in a publication, you can do so with high-resolution
            % Uncomment the following line to save the figure
            % print('combined_plots','-dpng','-r600'); % Saves the figure as a PNG with 600 DPI
            
            %%%%%%%%%%%%% Plotting Individually each plot %%%%%%%%%%%%%%%%
            % 
            % % Plotting the data
            % figure; % Open a new figure window
            % plot(data_table.Time_Corrected_Displacement, data_table.Displacement_Corrected_Displacement, 'LineWidth', 2);
            % 
            % % Setting the axes labels with LaTeX interpreter for symbols
            % xlabel('Time, s', 'Interpreter', 'latex', 'FontSize', 18);
            % ylabel('Displacement, $\mu$m', 'Interpreter', 'latex', 'FontSize', 18);
            % 
            % % Beautifications for nice visualization
            % set(gca, 'Box', 'on'); % Adding a box around the plot
            % set(gca, 'GridLineStyle', '-'); % Setting grid line style
            % grid on; % Adding a grid to the plot for easier readability of values
            % set(gcf, 'Color', 'w'); % Setting the figure background to white
            % set(gca, 'LineWidth', 1); % Making the axes lines thicker for visibility
            % set(gca, 'FontName', 'Times New Roman'); % Set the font to Times New Roman for publication standards
            % set(gca, 'FontSize', 14); % Setting a larger font size for the axes
            % set(gca, 'TickDir', 'in'); % Setting the tick direction to out
            % 
            % % If you want to save this figure for inclusion in a publication, you can do so with high resolution
            % % Uncomment the following line to save the figure
            % % print('myplot_time_vs_displacement','-dpng','-r600'); % Saves the figure as a PNG with 600 DPI
            % 
            % 
            % 
            % % Plotting the data
            % figure; % Open a new figure window
            % plot(data_table.Time_Corrected_Displacement, data_table.Load_Corrected_Load, 'LineWidth', 2);
            % 
            % % Setting the axes labels with LaTeX interpreter for symbols
            % xlabel('Time, s', 'Interpreter', 'latex', 'FontSize', 18);
            % ylabel('Load, mN', 'Interpreter', 'latex', 'FontSize', 18);
            % 
            % % Beautifications for nice visualization
            % set(gca, 'Box', 'on'); % Adding a box around the plot
            % set(gca, 'GridLineStyle', '-'); % Setting grid line style
            % grid on; % Adding a grid to the plot for easier readability of values
            % set(gcf, 'Color', 'w'); % Setting the figure background to white
            % set(gca, 'LineWidth', 1); % Making the axes lines thicker for visibility
            % set(gca, 'FontName', 'Times New Roman'); % Set the font to Times New Roman for publication standards
            % set(gca, 'FontSize', 14); % Setting a larger font size for the axes
            % set(gca, 'TickDir', 'in'); % Setting the tick direction to in
            % 
            % % If you want to save this figure for inclusion in a publication, you can do so with high resolution
            % % Uncomment the following line to save the figure
            % % print('myplot_time_vs_load','-dpng','-r600'); % Saves the figure as a PNG with 600 DPI
            % 
            % 
            % 
            % % Plotting the data
            % figure; % Open a new figure window
            % plot(data_table.Displacement_Corrected_Displacement, data_table.Load_Corrected_Load, 'LineWidth', 2);
            % 
            % % Setting the axes labels with LaTeX interpreter for symbols
            % xlabel('Displacement, $\mu$m', 'Interpreter', 'latex', 'FontSize', 18);
            % ylabel('Load, mN', 'Interpreter', 'latex', 'FontSize', 18);
            % 
            % % Beautifications for nice visualization
            % set(gca, 'Box', 'on'); % Adding a box around the plot
            % set(gca, 'GridLineStyle', '-'); % Setting grid line style
            % grid on; % Adding a grid to the plot for easier readability of values
            % set(gcf, 'Color', 'w'); % Setting the figure background to white
            % set(gca, 'LineWidth', 1); % Making the axes lines thicker for visibility
            % set(gca, 'FontName', 'Times New Roman'); % Set the font to Times New Roman for publication standards
            % set(gca, 'FontSize', 14); % Setting a larger font size for the axes
            % set(gca, 'TickDir', 'in'); % Setting the tick direction to out
            % 
            % % If you want to save this figure for inclusion in a publication, you can do so with high resolution
            % % Uncomment the following line to save the figure
            % % print('myplot','-dpng','-r600'); % Saves the figure as a PNG with 600 DPI
            
            %% Saving Figures and Processed Excel files
                new_filename = sprintf('%s_processed.xlsx', name);
                name_of_figure_plot = sprintf('%s_Plots.png', name);
                % Define the custom headers for the table
                customHeaders = {'Time(s)_Corrected_Displacement', 'Displacement(micro_meter)_Corrected_Displacement', ...
                             'Time(s)_Corrected_Load', 'Load(mN)_Corrected_Load', 'Raw_Time(s)_Raw_Displacement', ...
                             'Displacement(micro_meter)_Raw_Displacement', 'Raw_Time(s)_Raw_Load', 'Load(mN)_Raw_Load'};
                % Set the VariableNames property of the table to the custom headers
                data_table.Properties.VariableNames = customHeaders;
                writetable(data_table, new_filename);
                print(name_of_figure_plot,'-dpng','-r600')
                close all;
                % You might want to clear variables that will be reused in the next iteration
                clear raw_data processed_data data_table;
end

% At this point, each .xlsx file has been processed and saved with '_processed' suffix




%% Delete some patterned Files

%%%%%%%% Deleting _processed.xlsx Files %%%%%%%%%%
% % Get a list of all files in the current directory with names that end with '_processed.xlsx'
% 
% files = dir('*_processed.xlsx');
% % Loop over the files and delete each one
% for i = 1:length(files)
%     delete(files(i).name);
% end
% 
% %%%%%%%% Deleting *.png Files %%%%%%%%%%
% % % Get a list of all files in the current directory with names that end with '_processed.xlsx'
% 
% files = dir('*.png');
% 
% % Loop over the files and delete each one
% for i = 1:length(files)
%     delete(files(i).name);
% end



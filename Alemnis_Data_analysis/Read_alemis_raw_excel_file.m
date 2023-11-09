% Read the Excel file. This function call will get the data into a cell array.
% xlsread is used here without the first two output arguments which are for numeric data and text data separately.
[~, ~, raw_data] = xlsread('1.xlsx');

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

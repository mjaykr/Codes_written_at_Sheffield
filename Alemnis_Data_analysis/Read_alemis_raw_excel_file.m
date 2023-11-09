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

%% Removing the Rows having Positive or Zero time before negative time
% Find the index of the first negative value in the 'Time_Corrected_Displacement' column
is_negative = data_table.Time_Corrected_Displacement < 0;

% Find the first occurrence of a negative value
first_negative_index = find(is_negative, 1, 'first');

% Check if there is at least one negative value
if ~isempty(first_negative_index)
    % Check if the first value is non-negative and if there are positive values before the first negative value
    if first_negative_index > 1 && data_table.Time_Corrected_Displacement(1) >= 0
        % Remove the rows from the start until the first negative value (excluding the first negative value)
        data_table(1:first_negative_index-1, :) = [];
    end
end

% If there are no negative values, no rows are removed.



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

%% Unit Conversions from SI to the relevant unit for Plotting

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







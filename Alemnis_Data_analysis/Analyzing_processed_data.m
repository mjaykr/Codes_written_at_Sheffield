% MATLAB Script: Correct Experiment Start Considering Both Load and Displacement

% Define the filename
filename = dir('*_processed.xlsx'); % Finds any file with *_processed.xlsx
if isempty(filename)
    error('No *_processed.xlsx file found in the current folder.');
end

% Read the Excel file
filename = filename(1).name; % In case of multiple matches, take the first file
disp(['Reading data from file: ', filename]);

% Read the required columns
data = readtable(filename);

% Check if the necessary columns exist
requiredColumns = {'Time', 'Displacement', 'Load'};
if width(data) < 4
    error('The file does not have at least 4 columns.');
end

% Extract the necessary columns (assuming column indices: 1, 2, and 4)
time = data{:, 1};           % First column (Time)
displacement = data{:, 2};   % Second column (Displacement)
load = data{:, 4};           % Fourth column (Load)

% Determine the start of the experiment
positiveDisplacementIdx = find(displacement > 0, 1); % First positive displacement
positiveLoadIdx = find(load > 0, 1);                 % First positive load

if isempty(positiveDisplacementIdx) || isempty(positiveLoadIdx)
    error('No positive values found in Displacement or Load columns.');
end

% Ensure experiment starts when both Displacement and Load have positive values
startIdx = max(positiveDisplacementIdx, positiveLoadIdx);

% Find the point where displacement remains positive and load starts applying
while startIdx < length(displacement) && ...
      (any(displacement(startIdx:end) < 0) || any(load(startIdx:end) <= 0))
    startIdx = startIdx + 1;
end

if startIdx > length(displacement)
    error('Valid start point not found where both Displacement and Load are consistently positive.');
end

% Adjust time to start from zero
timeOffset = time(startIdx);
time = time - timeOffset;

% Adjust displacement to start from zero at the new start point
displacementOffset = displacement(startIdx);
displacement = displacement - displacementOffset;

% Adjust load to start from zero at the new start point
loadOffset = load(startIdx);
load = load - loadOffset;

% Truncate data to start from the new start index
time = time(startIdx:end);
displacement = displacement(startIdx:end);
load = load(startIdx:end);

% Create a corrected table with the desired column arrangement
correctedTable = table(time, load, time, displacement, displacement, load, ...
    'VariableNames', {'Time1', 'Load', 'Time2', 'Displacement1', 'Displacement2', 'Load2'});

% Display the corrected table
disp('Corrected Table:');
disp(correctedTable);

% Optionally, save the corrected data to a new Excel file
outputFilename = 'Corrected_Data_For_Plot.xlsx';
writetable(correctedTable, outputFilename);
disp(['The corrected table has been saved to ', outputFilename]);

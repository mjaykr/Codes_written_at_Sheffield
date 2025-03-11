% MATLAB Script to create folders based on .xlsx filenames and move files
% The operation is performed in the same directory where the script is run.

clc;
clear;

% Get the current directory (where the script is located)
sourceDir = pwd;

% Get all .xlsx files in the directory
filePattern = fullfile(sourceDir, '*.xlsx');
xlsxFiles = dir(filePattern);

% Check if there are any .xlsx files
if isempty(xlsxFiles)
    error('No .xlsx files found in the current directory.');
end

% Loop through each .xlsx file
for i = 1:length(xlsxFiles)
    % Get the current file's name and path
    currentFile = xlsxFiles(i).name; % Filename with extension
    [~, nameWithoutExt, ~] = fileparts(currentFile); % Filename without extension
    
    % Construct the folder name and paths
    folderName = fullfile(sourceDir, nameWithoutExt); % Folder name matches file name
    sourceFilePath = fullfile(sourceDir, currentFile); % Full path to the source file
    destinationFilePath = fullfile(folderName, currentFile); % Full path to the destination file
    
    % Create the folder if it doesn't exist
    if ~exist(folderName, 'dir')
        mkdir(folderName);
        disp(['Created folder: ', folderName]);
    end
    
    % Move the file to the folder
    try
        movefile(sourceFilePath, destinationFilePath);
        disp(['Moved file: ', currentFile, ' to folder: ', nameWithoutExt]);
    catch
        warning('Failed to move file "%s". Skipping...', currentFile);
    end
end

disp('Process completed.');

% Get the current working directory as the source folder
sourceFolder = pwd;

% Extract the parent directory (one folder up) for excelDestFolder
[parentDir, ~] = fileparts(sourceFolder);

% Define the destination folder for Excel files
excelDestFolder = fullfile(parentDir, 'Processed Excel Files');

% Define the destination folder for PNG files
pngDestFolder = fullfile(parentDir, 'Time Load Displacement Plots');

% Create the destination folders if they don't exist
if ~exist(excelDestFolder, 'dir')
    mkdir(excelDestFolder);
end

if ~exist(pngDestFolder, 'dir')
    mkdir(pngDestFolder);
end

% Move Excel files with the pattern *_processed.xlsx
excelPattern = fullfile(sourceFolder, '*_processed.xlsx');
excelFiles = dir(excelPattern);
for i = 1:numel(excelFiles)
    sourceFile = fullfile(sourceFolder, excelFiles(i).name);
    destinationFile = fullfile(excelDestFolder, excelFiles(i).name);
    movefile(sourceFile, destinationFile);
end

% Move PNG files with the pattern *.png
pngPattern = fullfile(sourceFolder, '*.png');
pngFiles = dir(pngPattern);
for i = 1:numel(pngFiles)
    sourceFile = fullfile(sourceFolder, pngFiles(i).name);
    destinationFile = fullfile(pngDestFolder, pngFiles(i).name);
    movefile(sourceFile, destinationFile);
end

disp('Files have been moved.');

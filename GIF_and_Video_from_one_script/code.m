% Combined MATLAB script to create both GIF and video from .tif files

% Get the current folder path and name
currentDirectory = pwd;
[~, folderName, ~] = fileparts(currentDirectory);

% Get a list of all files with .tif extension in the current directory
files = dir('*.tif');
numfiles = length(files);

% Check if there are any .tif files in the folder
if numfiles == 0
    error('No .tif files found in the current folder.');
end

% Read the first image to get its width
[imgFirst, cmapFirst] = imread(files(1).name);

% Convert to RGB if needed
if size(imgFirst, 3) ~= 3
    if isempty(cmapFirst)
        imgFirst = repmat(imgFirst, [1, 1, 3]);
    else
        imgFirst = ind2rgb(imgFirst, cmapFirst);
    end
end
imgFirst = im2uint8(imgFirst); % ensure uint8

% Use first image's width as targetWidth
targetWidth = size(imgFirst, 2);
aspectRatio = size(imgFirst, 2) / size(imgFirst, 1);
targetHeight = round(targetWidth / aspectRatio);

% Ensure both width and height are even
if mod(targetWidth, 2) ~= 0
    targetWidth = targetWidth + 1;
end
if mod(targetHeight, 2) ~= 0
    targetHeight = targetHeight + 1;
end

% Set the name for the GIF file
gifName = fullfile(pwd, [folderName, '.gif']);

% Frame rate setting
frameRate = 10;

% Create VideoWriter object
outputVideo = VideoWriter(fullfile(currentDirectory, folderName), 'MPEG-4');
outputVideo.FrameRate = frameRate;
open(outputVideo);

% Loop through each file
for k = 1:numfiles
    [img, cmap] = imread(files(k).name);
    
    if size(img, 3) ~= 3
        if isempty(cmap)
            img = repmat(img, [1, 1, 3]);
        else
            img = ind2rgb(img, cmap);
        end
    end
    
    if ~isa(img, 'uint8')
        img = im2uint8(img);
    end
    
    resizedImg = imresize(img, [targetHeight, targetWidth]);
    
    [A, map] = rgb2ind(resizedImg, 256);
    
    if k == 1
        imwrite(A, map, gifName, 'gif', 'LoopCount', Inf, 'DelayTime', 1/frameRate);
    else
        imwrite(A, map, gifName, 'gif', 'WriteMode', 'append', 'DelayTime', 1/frameRate);
    end
    
    writeVideo(outputVideo, resizedImg);
end

close(outputVideo);

disp(['Animated GIF saved as ', gifName]);
disp(['Video created successfully: ', fullfile(currentDirectory, [folderName, '.mp4'])]);

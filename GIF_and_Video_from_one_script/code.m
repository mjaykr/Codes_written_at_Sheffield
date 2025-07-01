% Combined MATLAB script to create both GIF and video from .tif files

% Get the current folder path and name
currentDirectory = pwd;
[~, folderName, ~] = fileparts(currentDirectory);

% Get a list of all files with .tif extension in the current directory
files = dir('*.tif');
numfiles = length(files);

if numfiles == 0
    error('No .tif files found in the current folder.');
end

% Read the first image to define dimensions
[imgFirst, cmapFirst] = imread(files(1).name);
if size(imgFirst, 3) ~= 3
    if isempty(cmapFirst)
        imgFirst = repmat(imgFirst, [1, 1, 3]);
    else
        imgFirst = ind2rgb(imgFirst, cmapFirst);
    end
end
imgFirst = im2uint8(imgFirst);

targetWidth = size(imgFirst, 2);
aspectRatio = size(imgFirst, 2) / size(imgFirst, 1);
targetHeight = round(targetWidth / aspectRatio);
if mod(targetWidth,2) ~= 0, targetWidth = targetWidth + 1; end
if mod(targetHeight,2) ~= 0, targetHeight = targetHeight + 1; end

% Set file names
gifName = fullfile(pwd, [folderName, '.gif']);
compressedGifName = fullfile(pwd, [folderName, '_compressed.gif']);
videoName = fullfile(pwd, [folderName, '.mp4']);

% Frame rate and delay time
frameRate = 3;
delayTime = 1 / frameRate;

% Initialize video writer
outputVideo = VideoWriter(videoName, 'MPEG-4');
outputVideo.FrameRate = frameRate;
open(outputVideo);

% Store frames in cell for re-use if needed
frames = cell(1, numfiles);

% Loop through each frame
for k = 1:numfiles
    [img, cmap] = imread(files(k).name);

    if size(img, 3) ~= 3
        if isempty(cmap)
            img = repmat(img, [1, 1, 3]);
        else
            img = ind2rgb(img, cmap);
        end
    end
    img = im2uint8(img);
    resizedImg = imresize(img, [targetHeight, targetWidth]);
    frames{k} = resizedImg;

    % Write to video
    writeVideo(outputVideo, resizedImg);
end
close(outputVideo);

% --------- Write original GIF at full resolution ---------
for k = 1:numfiles
    [A, map] = rgb2ind(frames{k}, 256);
    if k == 1
        imwrite(A, map, gifName, 'gif', 'LoopCount', Inf, 'DelayTime', delayTime);
    else
        imwrite(A, map, gifName, 'gif', 'WriteMode', 'append', 'DelayTime', delayTime);
    end
end

% --------- Check size of GIF and write compressed version if needed ---------
gifInfo = dir(gifName);
if gifInfo.bytes > 30 * 1024 * 1024  % 30 MB in bytes
    disp('GIF exceeds 30 MB. Writing compressed version...');

    for k = 1:numfiles
        % Reduce resolution and color depth
        resizedSmall = imresize(frames{k}, 0.5);  % 50% scaling
        [A, map] = rgb2ind(resizedSmall, 64);     % 64-color palette

        if k == 1
            imwrite(A, map, compressedGifName, 'gif', 'LoopCount', Inf, 'DelayTime', delayTime);
        else
            imwrite(A, map, compressedGifName, 'gif', 'WriteMode', 'append', 'DelayTime', delayTime);
        end
    end

    disp(['Compressed GIF saved as ', compressedGifName]);
else
    disp(['GIF saved as ', gifName]);
end

disp(['Video created successfully: ', videoName]);

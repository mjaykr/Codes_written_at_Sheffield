# README

## Overview

This MATLAB script processes all `.tif` files in the current directory to create both an animated GIF and a video file. The script resizes the images to a specified width while maintaining the aspect ratio, converts them to RGB if necessary, and then generates the GIF and video outputs.

## Requirements

- MATLAB (tested with R2021a, but should work with other versions)
- A directory containing `.tif` files

## Usage

1. Place the script in the directory containing your `.tif` files.
2. Run the script in MATLAB.

The script performs the following tasks:
- Reads all `.tif` files in the current directory.
- Resizes each image to a target width of 1080 pixels while maintaining the aspect ratio.
- Converts grayscale or indexed images to RGB.
- Creates an animated GIF from the resized images.
- Creates an MPEG-4 video from the resized images with a specified frame rate.

## Script Details

### Steps Performed by the Script

1. **Setup and Initialization**
   - Get the current directory path and name.
   - List all `.tif` files in the directory.
   - Check if there are any `.tif` files; if none are found, the script will terminate with an error.

2. **Define Parameters**
   - Set the target width for resizing images (`1080` pixels).
   - Define the output names for the GIF and video files.
   - Set the frame rate for the video (`3` frames per second).

3. **Process Each Image**
   - Loop through each `.tif` file:
     - Read the image.
     - Convert the image to RGB if it is grayscale or indexed.
     - Resize the image to the target width while maintaining the aspect ratio.
     - Convert the resized image to indexed color for the GIF.

4. **Create GIF**
   - Write the resized images to the GIF file, setting a loop count for infinite looping and a delay time between frames.

5. **Create Video**
   - Open a video writer object.
   - Write each resized image to the video file.
   - Close the video writer object.

6. **Output**
   - Display messages indicating the successful creation of the GIF and video files.

### Code

```matlab
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

% Define the target width
targetWidth = 1080;

% Set the name for the GIF file
gifName = fullfile(pwd, [folderName, '.gif']);

% Ask user for frame rate
% frameRate = input('Enter the frame rate for the video: ');
frameRate = 3;

% Create a VideoWriter object
outputVideo = VideoWriter(fullfile(currentDirectory, folderName), 'MPEG-4');
outputVideo.FrameRate = frameRate;

% Open the video file for writing
open(outputVideo);

% Loop through each file, resize it, and convert it
for k = 1:numfiles
    % Read the image
    [img, cmap] = imread(files(k).name);
    
    % Check if the image is grayscale or indexed and convert to RGB if necessary
    if size(img, 3) ~= 3
        if isempty(cmap)  % Image is grayscale
            img = repmat(img, [1, 1, 3]);
        else  % Image is indexed
            img = ind2rgb(img, cmap);
        end
    end
    
    % Ensure the image is uint8
    if ~isa(img, 'uint8')
        img = im2uint8(img);
    end
    
    % Calculate the new height to maintain the aspect ratio
    aspectRatio = size(img, 2) / size(img, 1);
    newHeight = round(targetWidth / aspectRatio);
    
    % Resize the image
    resizedImg = imresize(img, [newHeight, targetWidth]);
    
    % Convert the image to indexed color
    [A, map] = rgb2ind(resizedImg, 256);
    
    % Write the frame to the GIF file
    if k == 1
        imwrite(A, map, gifName, 'gif', 'LoopCount', Inf, 'DelayTime', 0.1);
    else
        imwrite(A, map, gifName, 'gif', 'WriteMode', 'append', 'DelayTime', 0.1);
    end
    
    % Write the resized image to the video
    writeVideo(outputVideo, resizedImg);
end

% Close the video file
close(outputVideo);

disp(['Animated GIF saved as ', gifName]);
disp(['Video created successfully: ', fullfile(currentDirectory, [folderName, '.mp4'])]);
```

## Notes

- Ensure that the directory contains `.tif` files before running the script.
- You can modify the `targetWidth` variable to resize images to a different width.
- The frame rate for the video can be adjusted by changing the `frameRate` variable.
- The GIF will have a delay time of `0.1` seconds between frames; this can be modified in the `imwrite` function call.

## Contact

For any questions or issues, please contact Dr. Mirtunjay Kumar at mjay@hotmail.com.

% Get a list of all files with .tif extension in the current directory
currentDirectory = pwd;
[path, folderName, ~] = fileparts(currentDirectory);
files = dir('*.tif');
numfiles = length(files);

% Define the target width
targetWidth = 1080;

% Set the name for the GIF file
gifName = fullfile(pwd, [folderName,'.gif']);

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
end

disp(['Animated GIF saved as ', gifName]);

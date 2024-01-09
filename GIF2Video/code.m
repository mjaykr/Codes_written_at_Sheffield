% Specify the folder path where your GIF files are located
folderPath = './'; % Replace with your folder path

% Get a list of all files in the folder with .gif extension
gifFiles = dir(fullfile(folderPath, '*.gif'));

% Loop through each GIF file
for k = 1:length(gifFiles)
    % Full path of the GIF file
    gifFilePath = fullfile(folderPath, gifFiles(k).name);
    
    % Read GIF file
    [gifImage, gifMap] = imread(gifFilePath, 'Frames', 'all');
    
    % Create a VideoWriter object for the output video
    % Change 'MPEG-4' to 'Motion JPEG AVI' for MPG format if needed
    [~, name, ~] = fileparts(gifFiles(k).name);
    outputVideoPath = fullfile(folderPath, [name '.mp4']);
    v = VideoWriter(outputVideoPath, 'MPEG-4');
    
    % Open the video file
    open(v);
    
    % Write each frame of the GIF to the video
    for frame = 1:size(gifImage, 4)
        % Convert indexed image to truecolor for video
        rgbImage = ind2rgb(gifImage(:, :, 1, frame), gifMap);
        
        % Write the frame to the video
        writeVideo(v, rgbImage);
    end
    
    % Close the video file
    close(v);
end

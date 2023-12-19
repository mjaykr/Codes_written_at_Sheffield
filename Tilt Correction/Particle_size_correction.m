% Parameters
apparentSize = 100; % Replace with the measured size from the SEM image
tiltAngleDegrees = 70; % Tilt angle in degrees

% Convert the tilt angle from degrees to radians
tiltAngleRadians = deg2rad(tiltAngleDegrees);

% Calculate the actual size of the particle
% Assuming the major axis of the particle is perpendicular to the tilt direction
actualSize = apparentSize / cos(tiltAngleRadians);

% Display the result
fprintf('The actual size of the particle is %.2f \x3BCm\n', actualSize);

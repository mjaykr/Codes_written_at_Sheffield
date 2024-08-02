clear
% Read the data from an Excel file in the same folder
[filename, filepath] = uigetfile('*.xlsx;*.xls', 'Select the Excel data file');
if isequal(filename, 0) || isequal(filepath, 0)
    disp('File selection cancelled');
    return;
end

% Read the Excel file
[~, ~, raw_data] = xlsread(fullfile(filepath, filename));

% Extract header and data
header = raw_data(1,:);
data = raw_data(2:end,:);

% Generate variable names from the header
variables = struct();
for i = 1:length(header)
    variable_name = matlab.lang.makeValidName(header{i});
    if isfloat(data{2,i})
        variables.(variable_name) = cell2mat(data(:,i));
    else
        variables.(variable_name) = data(:,i);
    end
end

% Find the column names for State of cycle, Orientation, and Particle Size
state_of_cycle_col = find(strcmpi(header, 'State of cycle'));
orientation_col = find(strcmpi(header, 'Orientation'));
particle_size_col = find(strcmpi(header, 'Particle Size'));

if isempty(state_of_cycle_col) || isempty(orientation_col)
    error('Could not find "State of cycle" or "Orientation" columns');
end

% Get unique State of cycle categories in the order they appear
state_of_cycle_data = variables.(matlab.lang.makeValidName(header{state_of_cycle_col}));
[unique_states, ~, state_indices] = unique(state_of_cycle_data, 'stable');

% Ask user which property to plot
properties = header(~ismember(header, {'State of cycle', 'Orientation', 'Particle Size'}));
disp('Available properties:');
for i = 1:length(properties)
    disp([num2str(i), '. ', properties{i}]);
end
property_index = input('Enter the number of the property you want to plot: ');

if property_index < 1 || property_index > length(properties)
    error('Invalid property selection');
end

selected_property = properties{property_index};
property_data = variables.(matlab.lang.makeValidName(selected_property));

% Ask user if they want to annotate Particle Size
annotate_particle_size = input('Do you want to annotate the Particle Size next to each scatter point? (Y/N): ', 's');

% Create figure
figure('Position', [100, 100, 800, 600]);
hold on;

% Define markers and colors for each orientation
markers = {'o', 's', '^'};  % Circle, Square, Triangle
colors = {'r', 'g', 'b'};   % Red, Green, Blue
unique_orientations = unique(variables.(matlab.lang.makeValidName(header{orientation_col})));

% Define offset
offset = 0.25;

% Plot data points
for i = 1:length(unique_orientations)
    orientation_type = unique_orientations{i};
    idx = strcmp(variables.(matlab.lang.makeValidName(header{orientation_col})), orientation_type);
    
    if strcmpi(orientation_type, 'Vertical')
        label = 'Parallel to Basal Plane';
    elseif strcmpi(orientation_type, 'Horizontal')
        label = 'Perpendicular to Basal Plane';
    else
        label = orientation_type;
    end
    
    scatter(state_indices(idx) + (rand(sum(idx),1)-0.5)*0.1, property_data(idx), 80, markers{i}, ...
        'filled', 'DisplayName', label, 'MarkerEdgeColor', 'k', 'MarkerFaceColor', colors{i});
    
    % Annotate Particle Size if user opted for it
    if strcmpi(annotate_particle_size, 'Y') && ~isempty(particle_size_col)
        particle_size_data = variables.(matlab.lang.makeValidName(header{particle_size_col}));
        text(state_indices(idx) + (rand(sum(idx),1)-0.5)*0.1, property_data(idx), ...
            cellstr(num2str(particle_size_data(idx))), 'FontSize', 12, 'HorizontalAlignment', 'center', 'VerticalAlignment', 'bottom');
    end
end

% Set x-axis ticks and labels
xticks(1:length(unique_states));
xticklabels(unique_states);
xlim([1-offset, length(unique_states)+offset]);

% Set axis labels and title
xlabel('State of Cycle', 'FontSize', 22);
ylabel([selected_property, ' ($$\mathrm{mN}$$)'], 'FontSize', 22, 'Interpreter', 'latex');
title('State of Cycle vs Selected Property', 'FontSize', 24);

% Add legend
legend('Location', 'best', 'FontSize', 18);

% Remove grid
grid off;

% Adjust tick properties
ax = gca;
ax.FontSize = 18;
ax.TickDir = 'in';
ax.TickLength = [0.02 0.02];
ax.XAxis.TickLength = [0.02 0.02];
ax.YAxis.TickLength = [0.02 0.02];

% Add top and right axes
ax.Box = 'on';
ax.XAxis.MinorTick = 'on';
ax.YAxis.MinorTick = 'on';

% Adjust figure size for better visibility
set(gcf, 'Position', [100, 100, 800, 600]);

% Tighten the plot
tight = get(gca, 'TightInset');
set(gca, 'Position', [tight(1)+0.02, tight(2)+0.02, 1-tight(1)-tight(3)-0.05, 1-tight(2)-tight(4)-0.05]);

hold off;

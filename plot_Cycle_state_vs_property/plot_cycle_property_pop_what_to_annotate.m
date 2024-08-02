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

% Find the column names for State of cycle and Orientation
state_of_cycle_col = find(strcmpi(header, 'State of cycle'));
orientation_col = find(strcmpi(header, 'Orientation'));

if isempty(state_of_cycle_col) || isempty(orientation_col)
    error('Could not find "State of cycle" or "Orientation" columns');
end

% Get unique State of cycle categories in the order they appear
state_of_cycle_data = variables.(matlab.lang.makeValidName(header{state_of_cycle_col}));
[unique_states, ~, state_indices] = unique(state_of_cycle_data, 'stable');

% Ask user which property to plot
properties = header(~ismember(header, {'State of cycle', 'Orientation'}));
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

% Ask user which property to annotate
disp('Available properties for annotation:');
for i = 1:length(header)
    disp([num2str(i), '. ', header{i}]);
end
annotate_index = input('Enter the number of the property you want to annotate (or 0 for no annotation): ');

if annotate_index < 0 || annotate_index > length(header)
    error('Invalid annotation selection');
end

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
    
    x_data = state_indices(idx) + (rand(sum(idx),1)-0.5)*0.1;
    y_data = property_data(idx);
    
    scatter(x_data, y_data, 80, markers{i}, ...
        'filled', 'DisplayName', label, 'MarkerEdgeColor', 'k', 'MarkerFaceColor', colors{i});
    
    % Annotate selected property if user opted for it
    if annotate_index > 0
        annotate_data = variables.(matlab.lang.makeValidName(header{annotate_index}));
        annotate_values = annotate_data(idx);
        
        % Smart text placement
        text_handles = text(x_data, y_data, cellstr(num2str(annotate_values)), ...
            'FontSize', 10, 'HorizontalAlignment', 'left', 'VerticalAlignment', 'bottom');
        
        % Adjust text positions to avoid overlap
        adjust_text(text_handles, x_data, y_data, annotate_values);
    end
end

% Set x-axis ticks and labels
xticks(1:length(unique_states));
xticklabels(unique_states);
xlim([1-offset, length(unique_states)+offset]);

% Set axis labels and title
% xlabel('State of Cycle', 'FontSize', 22);
ylabel(selected_property, 'FontSize', 22);
% title('State of Cycle vs Selected Property', 'FontSize', 24);

% Add legend
legend('Location', 'northeast', 'FontSize', 18);

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

% Function to adjust text positions
function adjust_text(text_handles, x_data, y_data, values)
    n = length(text_handles);
    positions = [x_data, y_data];
    
    for i = 1:n
        current_pos = positions(i,:);
        other_pos = positions([1:i-1, i+1:end], :);
        
        % Calculate distances to other points
        distances = sqrt(sum((other_pos - current_pos).^2, 2));
        
        % Find close points
        close_idx = distances < 0.5;  % Adjust this threshold as needed
        
        if any(close_idx)
            % Calculate average position of close points
            avg_pos = mean(other_pos(close_idx, :), 1);
            
            % Move text away from the average position
            direction = current_pos - avg_pos;
            direction = direction / norm(direction);
            
            offset = 0.2 * direction;  % Adjust multiplier for more/less spacing
            new_pos = current_pos + offset;
            
            % Update text position
            set(text_handles(i), 'Position', [new_pos, 0]);
        end
    end
end

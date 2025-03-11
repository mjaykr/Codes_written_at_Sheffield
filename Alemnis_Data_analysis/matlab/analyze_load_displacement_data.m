% Find and read the CSV file containing 'displacement_vs_load' in its name
files = dir('*displacement_vs_load*.csv');
data = readtable(files(1).name);

% Extract displacement and load from the first and second columns
displacement = data{:,1}; % Displacement in micrometers
load = data{:,2};         % Load in millinewtons

% Stiffness calculation Range
start_linear_range = 0.1;
end_linear_range = 0.3;

% Create a plot of load vs. displacement
figure;
h1 = plot(displacement, load, 'b-', 'LineWidth', 1.5); % Handle for load-displacement curve
xlabel('Displacement (um)');
ylabel('Load (mN)');
title('Load-Displacement Curve');
grid on;
hold on;

% Find the maximum load and the corresponding displacement
[max_load, idx] = max(load);
displacement_at_max_load = displacement(idx);
fprintf('Maximum Load: %.2f mN\n', max_load);
fprintf('Displacement at Maximum Load: %.2f um\n', displacement_at_max_load);

% Calculate energy to fracture using trapezoidal integration up to maximum load
energy_to_fracture = trapz(displacement(1:idx), load(1:idx));
fprintf('Energy to Fracture: %.2f nJ\n', energy_to_fracture); % 1 mN*um = 1 nJ

% Calculate stiffness from the initial linear portion (30% to 60% of max_load)
idx_stiffness = find(load >= start_linear_range*max_load & load <= end_linear_range*max_load & (1:length(load))' <= idx);
if ~isempty(idx_stiffness)
    p = polyfit(displacement(idx_stiffness), load(idx_stiffness), 1);
    stiffness = p(1);
    fprintf('Stiffness: %.2f mN/um\n', stiffness);
    % Plot the stiffness range data points
    h2 = plot(displacement(idx_stiffness), load(idx_stiffness), 'g.', 'MarkerSize', 10);
else
    fprintf('Unable to calculate stiffness: insufficient data in the specified range.\n');
    h2 = []; % Empty handle if no stiffness range
    stiffness = NaN; % Set to NaN if not calculated
end

% Detect pop-in events (sudden increase in displacement with small load change)
delta_displacement = diff(displacement);
delta_load = diff(load);
threshold_dd = prctile(delta_displacement, 95); % 95th percentile for large displacement jumps
threshold_dl = prctile(abs(delta_load), 25);    % 25th percentile for small load changes
pop_in_indices = find(delta_displacement > threshold_dd & abs(delta_load) < threshold_dl);

% Detect pop-out events (sudden decrease in load with small displacement change)
threshold_dl_neg = prctile(delta_load, 5); % 5th percentile for large negative load changes
pop_out_indices = find(delta_load < threshold_dl_neg & abs(delta_displacement) < prctile(abs(delta_displacement), 25));

% Plot detected events only if they exist
if ~isempty(pop_in_indices)
    h3 = plot(displacement(pop_in_indices + 1), load(pop_in_indices + 1), 'ro', 'MarkerSize', 8, 'LineWidth', 2);
else
    h3 = [];
end
if ~isempty(pop_out_indices)
    h4 = plot(displacement(pop_out_indices + 1), load(pop_out_indices + 1), 'bo', 'MarkerSize', 8, 'LineWidth', 2);
else
    h4 = [];
end

% Mark the maximum load point on the plot
plot(displacement_at_max_load, max_load, 'k*', 'MarkerSize', 10, 'LineWidth', 2);

% Add text annotations for key results on the plot
text_x = 0.05 * (max(displacement) - min(displacement)) + min(displacement); % 5% from left
text_y_start = 0.9 * max(load); % Start near top of y-axis
text_spacing = 0.1 * max(load); % Vertical spacing between lines
text(text_x, text_y_start, sprintf('Maximum Load: %.2f mN', max_load), 'FontSize', 10);
text(text_x, text_y_start - text_spacing, sprintf('Displacement at Max Load: %.2f um', displacement_at_max_load), 'FontSize', 10);
text(text_x, text_y_start - 2*text_spacing, sprintf('Energy to Fracture: %.2f nJ', energy_to_fracture), 'FontSize', 10);
if ~isnan(stiffness)
    text(text_x, text_y_start - 3*text_spacing, sprintf('Stiffness: %.2f mN/um', stiffness), 'FontSize', 10);
else
    text(text_x, text_y_start - 3*text_spacing, 'Stiffness: Not Calculated', 'FontSize', 10);
end

% Dynamically create the legend based on plotted elements
legend_entries = {'Load-Displacement', 'Max Load Point'};
legend_handles = [h1, plot(displacement_at_max_load, max_load, 'k*', 'MarkerSize', 10, 'LineWidth', 2)]; % Replot for legend handle
if ~isempty(h2)
    legend_entries{end+1} = 'Stiffness Range';
    legend_handles(end+1) = h2;
end
if ~isempty(pop_in_indices)
    legend_entries{end+1} = 'Pop-in Events';
    legend_handles(end+1) = h3;
end
if ~isempty(pop_out_indices)
    legend_entries{end+1} = 'Pop-out Events';
    legend_handles(end+1) = h4;
end
legend(legend_handles, legend_entries, 'Location', 'best');
hold off;

% Save the plot as a PNG file
saveas(gcf, 'load_displacement_analysis.png');

% Print pop-in event details
fprintf('Number of Pop-in Events: %d\n', length(pop_in_indices));
for i = 1:length(pop_in_indices)
    fprintf('Pop-in Event %d: Displacement = %.2f um, Load = %.2f mN\n', ...
        i, displacement(pop_in_indices(i)+1), load(pop_in_indices(i)+1));
end

% Print pop-out event details
fprintf('Number of Pop-out Events: %d\n', length(pop_out_indices));
for i = 1:length(pop_out_indices)
    fprintf('Pop-out Event %d: Displacement = %.2f um, Load = %.2f mN\n', ...
        i, displacement(pop_out_indices(i)+1), load(pop_out_indices(i)+1));
end

% Analyze post-fracture behavior
if idx < length(load)
    post_fracture_load = load(idx+1:end);
    min_post_load = min(post_fracture_load);
    idx_min = find(load == min_post_load, 1, 'first');
    displacement_at_min_load = displacement(idx_min);
    fprintf('Minimum Load After Fracture: %.2f mN\n', min_post_load);
    fprintf('Displacement at Minimum Load After Fracture: %.2f um\n', displacement_at_min_load);
else
    fprintf('No data available after fracture.\n');
end

% Print final displacement and load
fprintf('Final Displacement: %.2f um\n', displacement(end));
fprintf('Final Load: %.2f mN\n', load(end));

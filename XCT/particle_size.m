% Read particle sizes from the text file
filename = 'particle_sizes.txt'; % Specify the path if the file is not in the current directory
particleSizes = readmatrix(filename);

% Calculate the relative frequency of particle sizes
edges = linspace(min(particleSizes), max(particleSizes), 20); % Adjust the number of bins as needed
[counts, edges] = histcounts(particleSizes, edges);
relativeFrequency = counts / sum(counts);

% Calculate the center of each bin for plotting
binCenters = edges(1:end-1) + diff(edges)/2;

% Plotting for Scientific Publication
figure;
bar(binCenters, relativeFrequency, 'BarWidth', 1, 'FaceColor', [0 0.4470 0.7410]);

% Improve the aesthetics for publication
set(gca, 'FontSize', 14, 'LineWidth', 1.5); % Adjust font size and axis line width
xlabel('Particle Size', 'FontSize', 16);
ylabel('Relative Frequency', 'FontSize', 16);
title('Particle Size Distribution', 'FontSize', 18);

% Set tick direction to inside
set(gca, 'TickDir', 'in');

% Add grid lines and set axis limits tightly around the data
grid on;
axis tight;

% Optionally, set the figure size and background to white
set(gcf, 'Color', 'w', 'Units', 'Inches', 'Position', [0, 0, 6, 4.5]); % 6x4.5 inch figure

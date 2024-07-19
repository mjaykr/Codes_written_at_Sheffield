clear
load data_5K_1mu_dia.mat
x = data.X;
y = data.Y;

y1_values = input('Enter the array of y value of interest (y_array): ');
area_all = zeros(1,length(y1_values));

for i = 1:length(y1_values)

    y1 = y1_values(i);
    
    % Filter points where y <= y1
    indices = y <= y1;
    x_filtered = x(indices);
    y_filtered = y(indices);
    
    % Fit a curve to the data points (polynomial fit as an example)
    p = polyfit(x_filtered, y_filtered, 5); % 5th degree polynomial fit
    f = @(x_filtered) polyval(p, x_filtered);
    f_prime = @(x_filtered) polyval(polyder(p), x_filtered);
    
    % Define the integrand for the surface area of revolution
    integrand = @(x_filtered) 2 * pi * f(x_filtered) .* sqrt(1 + f_prime(x_filtered).^2);
    
    % Compute the integral for the surface area of revolution
    area = integral(integrand, min(x_filtered), max(x_filtered));
    
    % Display the calculated area
    disp(['Area of contact during indentation is: ', num2str(area), ' square micrometers']);
    area_all(i) = area;

end

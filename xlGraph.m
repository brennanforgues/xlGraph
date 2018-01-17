function xlGraph(varargin)

xl = actxserver('Excel.Application'); % excel application
xl.Visible = true; % set excel to be visible to the user
xlSheet = xl.Workbooks.Add.ActiveSheet; % active excel sheet

% define xl chart 
xlChart = xlSheet.ChartObjects.Add(100, 30, 400, 250).Chart; 

% Set chart type and title
xlChart.ChartType = 65;
xlChart.ChartType = 'xlXYScatterSmooth';
xlChart.HasTitle = true;
xlChart.ChartTitle.Text = 'Figure 1';

if mod(nargin, 2) ~= 0
    error('Must have even number of input arguments to xlGraph');
else
    seriesCount = 1; % keep track of the number of series' we are dealing with
    for i = 1:2:nargin
        % bind our series to the xl chart    
        xlChart.SeriesCollection.NewSeries;
        xlChart.SeriesCollection(seriesCount).XValue = varargin{i}; % x values on the series
        xlChart.SeriesCollection(seriesCount).Value = varargin{i + 1}; % y values on the series
        seriesCount = seriesCount + 1;
    end   
end

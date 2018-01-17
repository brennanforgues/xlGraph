function xlGraph(x,y,x2,y2,x3,y3,x4,y4,x5,y5)
% function takes in up to 5 curves and displays data on the same graph

xl = actxserver('Excel.Application'); % excel application
xlWs = xl.Workbooks; % excel worksheet 
xlW = xlWs.Add; % excel document 
xlS = xlW.ActiveSheet; % excel sheet
xl.Visible = 1;

L = int16.empty;

% define xl chart 
xlCO = xlS.ChartObjects.Add(100, 30, 400, 250); 
xlC = xlCO.Chart;

% Set chart type and title
xlC.ChartType = 1;
xlC.ChartType = 65;
xlC.ChartType = 'xlXYScatterSmooth';
xlC.HasTitle = true;
xlC.ChartTitle.Text = 'Figure 1: ';


if nargin > 10
    error('Too many input arguments to the xlGraph function');
end

% transpose arrays
if nargin == 10
    
     if size(x5,1) == 1
         x5 = x5';
     end
     if size(y5,1) == 1
         y5 = y5';
     end
     if size(x4,1) == 1
         x4 = x4';
     end
     if size(y4,1) == 1
         y4 = y4';
     end
     if size(x3,1) == 1
         x3 = x3';
     end
     if size(y3,1) == 1
         y3 = y3';
     end
     if size(x2,1) == 1
         x2 = x2';
     end
     if size(y2,1) == 1
         y2 = y2';
     end
     if size(x,1) == 1
         x = x';
     end
     if size(y,1) == 1
         y = y';
     end
     
    % Bind input arrays to xl range object
    xlS.Range('A1:B1000').Value = [x y];
    xlS.Range('C1:D1000').Value = [x2 y2];
    xlS.Range('E1:F1000').Value = [x3 y3];
    xlS.Range('G1:H1000').Value = [x4 y4];
    xlS.Range('I1:J1000').Value = [x5 y5];
    
    % Bind xl range object to xl chart object 
    xlC.SeriesCollection.NewSeries;
    xlC.SeriesCollection(1).Value = xlS.Range('B1:B1000');
    xlC.SeriesCollection(1).XValue = xlS.Range('A1:A1000');
    xlC.SeriesCollection.NewSeries;
    xlC.SeriesCollection(2).Value = xlS.Range('D1:D1000');
    xlC.SeriesCollection(2).XValue = xlS.Range('C1:C1000');
    xlC.SeriesCollection.NewSeries;
    xlC.SeriesCollection(3).Value = xlS.Range('F1:F1000');
    xlC.SeriesCollection(3).XValue = xlS.Range('E1:E1000');
    xlC.SeriesCollection.NewSeries;
    xlC.SeriesCollection(4).Value = xlS.Range('H1:H1000');
    xlC.SeriesCollection(4).XValue = xlS.Range('G1:G1000');
    xlC.SeriesCollection.NewSeries;
    xlC.SeriesCollection(5).Value = xlS.Range('J1:J1000');
    xlC.SeriesCollection(5).XValue = xlS.Range('I1:I1000');
    
elseif nargin == 9
     error('Must have even number of input arguments to xlGraph');

% transpose arrays
elseif nargin == 8
     x5 = L;
     y5 = L;
     
     
     if size(x4,1) == 1
         x4 = x4';
     end
     if size(y4,1) == 1
         y4 = y4';
     end
     if size(x3,1) == 1
         x3 = x3';
     end
     if size(y3,1) == 1
         y3 = y3';
     end
     if size(x2,1) == 1
         x2 = x2';
     end
     if size(y2,1) == 1
         y2 = y2';
     end
     if size(x,1) == 1
         x = x';
     end
     if size(y,1) == 1
         y = y';
     end
        
     % Bind input arrays to xl range object
     xlS.Range('A1:B1000').Value = [x y];
     xlS.Range('C1:D1000').Value = [x2 y2];
     xlS.Range('E1:F1000').Value = [x3 y3];
     xlS.Range('G1:H1000').Value = [x4 y4];
     
     % Bind xl range object to xl chart object
     xlC.SeriesCollection.NewSeries;
     xlC.SeriesCollection(1).Value = xlS.Range('B1:B1000');
     xlC.SeriesCollection(1).XValue = xlS.Range('A1:A1000');
     xlC.SeriesCollection.NewSeries;
     xlC.SeriesCollection(2).Value = xlS.Range('D1:D1000');
     xlC.SeriesCollection(2).XValue = xlS.Range('C1:C1000');
     xlC.SeriesCollection.NewSeries;
     xlC.SeriesCollection(3).Value = xlS.Range('F1:F1000');
     xlC.SeriesCollection(3).XValue = xlS.Range('E1:E1000');
     xlC.SeriesCollection.NewSeries;
     xlC.SeriesCollection(4).Value = xlS.Range('H1:H1000');
     xlC.SeriesCollection(4).XValue = xlS.Range('G1:G1000');


elseif nargin == 7
     error('Must have even number of input arguments to xlGraph');
     
% transpose arrays     
elseif nargin == 6
     x5 = L;
     y5 = L;
     x4 = L;
     y4 = L;
     
     
     if size(x3,1) == 1
         x3 = x3';
     end
     if size(y3,1) == 1
         y3 = y3';
     end
     if size(x2,1) == 1
         x2 = x2';
     end
     if size(y2,1) == 1
         y2 = y2';
     end
     if size(x,1) == 1
         x = x';
     end
     if size(y,1) == 1
         y = y';
     end
        
     % Bind input arrays to xl range object
     xlS.Range('A1:B1000').Value = [x y];
     xlS.Range('C1:D1000').Value = [x2 y2];
     xlS.Range('E1:F1000').Value = [x3 y3];
     
     % Bind xl range object to xl chart object
     xlC.SeriesCollection.NewSeries;
     xlC.SeriesCollection(1).Value = xlS.Range('B1:B1000');
     xlC.SeriesCollection(1).XValue = xlS.Range('A1:A1000');
     xlC.SeriesCollection.NewSeries;
     xlC.SeriesCollection(2).Value = xlS.Range('D1:D1000');
     xlC.SeriesCollection(2).XValue = xlS.Range('C1:C1000');
     xlC.SeriesCollection.NewSeries;
     xlC.SeriesCollection(3).Value = xlS.Range('F1:F1000');
     xlC.SeriesCollection(3).XValue = xlS.Range('E1:E1000');
elseif nargin == 5
     error('Must have even number of input arguments to xlGraph');  
     
% transpose arrays     
elseif nargin == 4
     x5 = L;
     y5 = L;
     x4 = L;
     y4 = L;
     x3 = L;
     y3 = L;
     
     if size(x2,1) == 1
         x2 = x2';
     end
     if size(y2,1) == 1
         y2 = y2';
     end
     if size(x,1) == 1
         x = x';
     end
     if size(y,1) == 1
         y = y';
     end
       
     % Bind input arrays to xl range object
     xlS.Range('A1:B1000').Value = [x y];
     xlS.Range('C1:D1000').Value = [x2 y2];
     
     % Bind xl range object to xl chart object
     xlC.SeriesCollection.NewSeries;
     xlC.SeriesCollection(1).Value = xlS.Range('B1:B1000');
     xlC.SeriesCollection(1).XValue = xlS.Range('A1:A1000');
     xlC.SeriesCollection.NewSeries;
     xlC.SeriesCollection(2).Value = xlS.Range('D1:D1000');
     xlC.SeriesCollection(2).XValue = xlS.Range('C1:C1000');
elseif nargin == 3
     error('Must have even number of input arguments to xlGraph');
     
% transpose arrays     
elseif nargin == 2
     x5 = L;
     y5 = L;
     x4 = L;
     y4 = L;
     x3 = L;
     y3 = L;
     x2 = L;
     y2 = L;
     
     
     if size(x,1) == 1
         x = x';
     end
     if size(y,1) == 1
         y = y';
     end
    
     
     % Bind input arrays to xl range object   
     xlS.Range('A1:B1000').Value = [x y];
     
     % Bind xl range object to xl chart object
     xlC.SeriesCollection.NewSeries;
     xlC.SeriesCollection(1).Value = xlS.Range('B1:B1000');
     xlC.SeriesCollection(1).XValue = xlS.Range('A1:A1000');
     
elseif nargin == 1
     error('Must have even number of input arguments to xlGraph');
     
end











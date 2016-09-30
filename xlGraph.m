function xlGraph(x,y,x2,y2,x3,y3,x4,y4,x5,y5)

e = actxserver('Excel.Application'); % e = word
eWs = e.Workbooks;
eW = eWs.Add; % eW = document 
eS = eW.ActiveSheet;
e.Visible = 1;



L = int16.empty;

eCO = eS.ChartObjects.Add(100, 30, 400, 250);
eC = eCO.Chart;



 
if nargin > 10
    error('Too many input arguments to the xlsplot function');
end

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
    eS.Range('A1:B1000').Value = [x y];
    eS.Range('C1:D1000').Value = [x2 y2];
    eS.Range('E1:F1000').Value = [x3 y3];
    eS.Range('G1:H1000').Value = [x4 y4];
    eS.Range('I1:J1000').Value = [x5 y5];
    
    eC.SeriesCollection.NewSeries;
    eC.SeriesCollection(1).Value = eS.Range('B1:B1000');
    eC.SeriesCollection(1).XValue = eS.Range('A1:A1000');
    eC.SeriesCollection.NewSeries;
    eC.SeriesCollection(2).Value = eS.Range('D1:D1000');
    eC.SeriesCollection(2).XValue = eS.Range('C1:C1000');
    eC.SeriesCollection.NewSeries;
    eC.SeriesCollection(3).Value = eS.Range('F1:F1000');
    eC.SeriesCollection(3).XValue = eS.Range('E1:E1000');
    eC.SeriesCollection.NewSeries;
    eC.SeriesCollection(4).Value = eS.Range('H1:H1000');
    eC.SeriesCollection(4).XValue = eS.Range('G1:G1000');
    eC.SeriesCollection.NewSeries;
    eC.SeriesCollection(5).Value = eS.Range('J1:J1000');
    eC.SeriesCollection(5).XValue = eS.Range('I1:I1000');
    
elseif nargin == 9
     error('Must have even number of input arguments to xlsplot');
     
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
        
     eS.Range('A1:B1000').Value = [x y];
     eS.Range('C1:D1000').Value = [x2 y2];
     eS.Range('E1:F1000').Value = [x3 y3];
     eS.Range('G1:H1000').Value = [x4 y4];
     
     eC.SeriesCollection.NewSeries;
     eC.SeriesCollection(1).Value = eS.Range('B1:B1000');
     eC.SeriesCollection(1).XValue = eS.Range('A1:A1000');
     eC.SeriesCollection.NewSeries;
     eC.SeriesCollection(2).Value = eS.Range('D1:D1000');
     eC.SeriesCollection(2).XValue = eS.Range('C1:C1000');
     eC.SeriesCollection.NewSeries;
     eC.SeriesCollection(3).Value = eS.Range('F1:F1000');
     eC.SeriesCollection(3).XValue = eS.Range('E1:E1000');
     eC.SeriesCollection.NewSeries;
     eC.SeriesCollection(4).Value = eS.Range('H1:H1000');
     eC.SeriesCollection(4).XValue = eS.Range('G1:G1000');
elseif nargin == 7
     error('Must have even number of input arguments to xlsplot');
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
        
     eS.Range('A1:B1000').Value = [x y];
     eS.Range('C1:D1000').Value = [x2 y2];
     eS.Range('E1:F1000').Value = [x3 y3];
     
     eC.SeriesCollection.NewSeries;
     eC.SeriesCollection(1).Value = eS.Range('B1:B1000');
     eC.SeriesCollection(1).XValue = eS.Range('A1:A1000');
     eC.SeriesCollection.NewSeries;
     eC.SeriesCollection(2).Value = eS.Range('D1:D1000');
     eC.SeriesCollection(2).XValue = eS.Range('C1:C1000');
     eC.SeriesCollection.NewSeries;
     eC.SeriesCollection(3).Value = eS.Range('F1:F1000');
     eC.SeriesCollection(3).XValue = eS.Range('E1:E1000');
elseif nargin == 5
     error('Must have even number of input arguments to xlsplot');  
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
        
     eS.Range('A1:B1000').Value = [x y];
     eS.Range('C1:D1000').Value = [x2 y2];
     
     eC.SeriesCollection.NewSeries;
     eC.SeriesCollection(1).Value = eS.Range('B1:B1000');
     eC.SeriesCollection(1).XValue = eS.Range('A1:A1000');
     eC.SeriesCollection.NewSeries;
     eC.SeriesCollection(2).Value = eS.Range('D1:D1000');
     eC.SeriesCollection(2).XValue = eS.Range('C1:C1000');
elseif nargin == 3
     error('Must have even number of input arguments to xlsplot');
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
    
     
        
     eS.Range('A1:B1000').Value = [x y];
     
     eC.SeriesCollection.NewSeries;
     eC.SeriesCollection(1).Value = eS.Range('B1:B1000');
     eC.SeriesCollection(1).XValue = eS.Range('A1:A1000');
     
elseif nargin == 1
     error('Must have even number of input arguments to xlsplot');
     
end


eCO.Chart.ChartType = 1;
eCO.Chart.ChartType = 65;
eCO.Chart.ChartType = 'xlXYScatterSmooth';

eCO.Chart.HasTitle = true;
eCO.Chart.ChartTitle.Text = 'Figure 1: ';


%eCO.Chart.Axes(xlCategory, xlPrimary).HasTitle = true;
%eCO.Chart.Axes(xlCategory, xlPrimary).AxisTitle.Text = 'X Axis';

%eCO.Chart.Axes(xlValue, xlPrimary).HasTitle = true;
%eCO.Chart.Axes(xlValue, xlPrimary).AxisTitle.Text = 'Y Axis';







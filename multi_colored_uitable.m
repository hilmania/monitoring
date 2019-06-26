% function multi_colored_uitable
% Create a figure
f = figure;
% Define the column names
colnames = {'first', 'second', 'third', 'fourth','fifth','sixth'};
% Create a cell array with HTML code
celldata = {colText('Red text', 'red'), ...
      colText('Blue text', 'blue'), ...
      colText('Green text', 'green'), ...
      colText('Purple text', 'purple'), ...
      colText('RGB Green','rgb(0,255,0)'),...
      colText('HEX Green','#00ff00')};
% Create a uitable with one row and 4 columns and colored cells
a = uitable(f, 'Data', celldata, ...
      'ColumnName', colnames, ...
      'Units', 'normalized', ...
      'Position', [0 0 1 1]);
% end

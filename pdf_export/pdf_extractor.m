clc; clear; close all;



folderPath = fileparts(mfilename('fullpath'));
path = folderPath+ "\data";
files = dir(fullfile(path, '*.pdf'));
patient_number = 1;

for n = 1:height(files)
    file = fullfile(path, files(n).name);
    str = extractFileText(file);
    filename = fopen('text.txt','w');
    fwrite(filename,str);
    
    raw = readlines_combined();
    %raw = readlines("text.txt","EmptyLineRule","skip","WhitespaceRule","trim");
    fclose(filename);
    
    table = [];
    
    i=1;
    column_count = 1;
    new = true;
    name = [];
    
    while i < length(raw)
        %oszlopok szama    
        if strcmp(raw(i), 'Course ID')
            while ~contains(raw(i), 'Achieved') && new 
            column_count = column_count+1;
            name = [name, raw(i)];
            i = i+1;
            end
                if new
                    name = [name, 'Achieved'];
                end
            new = false;
        end
        %Ha kulon van az ID
        if strcmp(raw(i), 'Course')
            while ~contains(raw(i), 'Achieved') && new 
            column_count = column_count+1;
            name = [name, raw(i)];
            i = i+1;
            end
                if new
                    name = [name, 'Achieved'];
                end
            new = false;

            column_count = column_count - 3;
        end

        
        % azonosito szam
        if contains(raw(i), 'Patient:')
            id = regexp(str, '\((\d+)\)', 'tokens', 'once');
    
    
        end
    
        if contains(raw(i), 'C1')
            new_row = [];
            for k = 1:column_count
                new_row = [new_row, raw(i)];
                i = i+1;
            end
            table = [table; new_row];
            i = i-1;
        end
    
    i = i+1;
    
    
    end

    %Pagebreak hibak javitasa plan id-ra
    
    for i = 2:length(table(:,2))
        if ~strcmp(table(i,2), table(i-1,2)) && strcmp(table(i-1,2), table(i+1,2))
            table(i,2) = table(i-1,2);
        end

    end

        
    %%Mertekegyseg szetvalasztasa

    table = [table(:,2),table(:,4),table(:,5),table(:,7),table(:,8)];
    
    table = separate_unit(table);



    %tipusok szetvalasztasa egymas melle

    column_count = 5;
    separated_table = separate_methods(table,column_count);

    % %Omlesztes
    % 
    % sum_table = strings(length(files(:,1))+2,length(separated_table(:,1)));
    % 
    %     if patient_number == 1
    %     %Fejlec
    %     sum_table(1,:) = separated_table(:,4)'; % structure name id
    %     sum_table(2,:) = separated_table(:,5)'; % objective
    %     end
    %     %
    %     sum_table(patient_number+2, :) = separated_table(:,7)'; % data _ RA
    %     patient_number = patient_number +1;
    % 


    table = array2table(separated_table);


    table = [table,array2table(zeros(size(table(:,1))))];
   
    

    
    file_name = id + '.xlsx';
    location = fullfile(folderPath, "\excel", file_name);

    if isfile(location)
        delete(location);
    end

    %Extra kezelesek levagasa
    filtered_table = [];

    for i = 1:width(table)-1
        line = table2cell(table(height(table),i));
        if ~ismissing(line{1})
            filtered_table = [filtered_table,table2cell(table(:,i))];
        end

    end

    table = cell2table(filtered_table);
    filtered_name = [name(2);name(4);name(5);name(7);name(8)];
    table.Properties.VariableNames(1:3*(column_count+1)) = [filtered_name;'Unit1'; filtered_name + '2'; 'Unit2'; filtered_name + '3' ;'Unit3'];
    

    %Tipus konverzio
    
    nums = [table.Actual,table.Actual2, table.Actual3];
    nums = double(nums);


    for k = 1:width(table)
    table.(k) = cellstr(table.(k));  % Convert each string column to cell array of char vectors
    end

    table.Actual= nums(:,1);
    table.Actual2 = nums(:,2);
    table.Actual3 = nums(:,3);

    table = rows2vars(table);

    

    writetable(table,location);

end

function raw = readlines_combined()
rawLines = readlines("text.txt","EmptyLineRule","read");

combinedLines = {};  % Store combined lines
accumulator = "";    % Accumulate lines without empty line separation

for i = 1:length(rawLines)
    line = rawLines(i);
    if strlength(strtrim(line)) == 0
        % Empty line found: store accumulated line if not empty
        if strlength(accumulator) > 0
            combinedLines{end+1} = accumulator; %#ok<SAGROW>
            accumulator = "";
        end
    else
        % Non-empty line: append (add a space between lines)
        if strlength(accumulator) > 0
            accumulator = accumulator + " " + strtrim(line);
        else
            accumulator = strtrim(line);
        end
    end
end
raw = string(combinedLines');

end

function table = separate_unit(table)
numericActual = zeros(length(table(:,1)), 1);
    unitColumn = strings(length(table(:,1)), 1);  % use string array instead of cell
    
    for i = 1:length(table(:,1))
        %if contains(table(i,4), ' ')
            number = str2double(regexp(table(i,4), '[\d.]+', 'match', 'once'));
            unit = strtrim(regexp(table(i,4), '[^\d\s.]+.*$', 'match', 'once'));
    
            numericActual(i) = number;
            unitColumn(i) = unit;
       % end
    end
    
    % Replace or add proper columns
    table(:,4) = numericActual;
    table(:,6) = unitColumn;
    
end

function separated_table = separate_methods(table,column_count)
separated_table = ["string"];
    column_count = column_count +1;
    count = 1;
    k = 1;
    for i = 2:length(table(:,1))
        if strcmp(table(i,1), table(i-1,1))
            separated_table(k, count*column_count-column_count+1:count*column_count) = table(i-1,:);
            k = k+1;
        else
            separated_table(k, count*column_count-column_count+1:count*column_count) = table(i-1,:);
            k= 1;
            count = count+1;            
        end
    end
    separated_table(k, count*column_count-column_count+1:count*column_count) = table(i,:);

end

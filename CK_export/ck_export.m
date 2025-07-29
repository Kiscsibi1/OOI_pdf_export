clc; clear; close all;

%testing pull

addpath('C:\Users\Gazdi\Documents\SynologyDrive\tanulmanyok\6.BME 2024-2025-2\Szakgyak\pdf_export');

pdf_extractor;

clc; clear; close all;


folderPath = fileparts(mfilename('fullpath'));
path = folderPath+ "\data";
files = dir(fullfile(path, '*.xlsx'));


for patient = 1:height(files)


    file = fullfile(path, files(patient).name);

if startsWith(files(patient).name, "~")
    continue
end

excel_raw = readcell(file,'Sheet','Sheet3');

excel_raw = cell2table(excel_raw);
data = [];


for i = 1:height(excel_raw)
    if strcmp(string(excel_raw.excel_raw2(i)),"Dx Vx Values")
        split_height = i;
        for j = i+2:height(excel_raw)
            if ismissing(string(excel_raw.excel_raw2(j)))
                end_height = j-1;
                break
            
            end

        end
        break
    end
end

raw_width = width(excel_raw);

excel_raw_VOI = table();
excel_raw_DXVX = table();
excel_raw_VOI = excel_raw(1:split_height-1,1:raw_width);
excel_raw_DXVX = excel_raw(split_height:end_height,1:raw_width);


% Fully missing columns removed
    for i = 1:raw_width
        column_name = "excel_raw_VOI" + ".excel_raw" + string(i);
        column = eval(column_name);
        column = string(column);
        tmp = sum(double(ismissing(column)));
        if sum(double(ismissing(column))) < length(column(:,1))
            data = [data, column];
        end
    end

    excel_raw_VOI = table();
    excel_raw_VOI = array2table(data);
    data = [];

    for i = 1:raw_width
        column_name = "excel_raw_DXVX" + ".excel_raw" + string(i);
        column = eval(column_name);
        column = string(column);
        tmp = ismissing(column);
        if sum(double(ismissing(column))) < length(column(:,1))
            data = [data, column];
        end
    end

    excel_raw_DXVX = table();
    excel_raw_DXVX = array2table(data);


%Reading normal_max for normalizing
normal_max = excel_raw_DXVX{3,3};
normal_max = str2double(regexp(normal_max, '[\d.]+', 'match', 'once'));

% Searching for the DXVX table
names = ["DVH";"Dose_cGy";"Dose_p";"Volume_cm3";"Volume_p"];

data_DXVX = excel_raw_DXVX(4:end,1:5);

data_DXVX.Properties.VariableNames = names;

for i = 2:5
data_DXVX.(i)= double(data_DXVX.(i));
end

% Searching for the VOI table
names = ["VOI"; "Volume"; "Mean"; "Max"];

for i = 11:height(excel_raw_VOI)
    if ismissing(string(excel_raw_VOI.data4(i)))
        end_height = i-1;
        break    
    end
end

data_VOI = excel_raw_VOI(11:end_height,[4,5,7,8]);

data_VOI.Properties.VariableNames = names;
for i = 2:4
data_VOI.(i)= double(data_VOI.(i));
end


n = size(data_VOI, 1);
struct_VOI = struct();

for i = 1:n
    name = matlab.lang.makeValidName(erase(lower(data_VOI{i,1}),[" " , "_", "-", "_"]));  % Ensure it's a valid field name
    struct_VOI.(name).Volume = data_VOI.Volume(i);
    struct_VOI.(name).Mean   = data_VOI.Mean(i);
    struct_VOI.(name).Max    = data_VOI.Max(i);
end

n = size(data_DXVX, 1);
struct_DXVX = struct();

for i = 1:n
    name = matlab.lang.makeValidName(erase(lower(data_DXVX{i,1}),[" " , "_", "-", "_"]));
    if isfield(struct_DXVX, name)
        struct_DXVX.(name).Dose_cGy = [struct_DXVX.(name).Dose_cGy; data_DXVX.Dose_cGy(i)];
        struct_DXVX.(name).Dose_p = [struct_DXVX.(name).Dose_p; data_DXVX.Dose_p(i)];
        struct_DXVX.(name).Volume_cm3 = [struct_DXVX.(name).Volume_cm3; data_DXVX.Volume_cm3(i)];
        struct_DXVX.(name).Volume_p = [struct_DXVX.(name).Volume_p;data_DXVX.Volume_p(i)];
    
    else
    struct_DXVX.(name).Dose_cGy = data_DXVX.Dose_cGy(i);
    struct_DXVX.(name).Dose_p = data_DXVX.Dose_p(i);
    struct_DXVX.(name).Volume_cm3 = data_DXVX.Volume_cm3(i);
    struct_DXVX.(name).Volume_p = data_DXVX.Volume_p(i);
    end

end

connections= readtable("Connections.xlsx");

PlanSetup_ID4 = cell(height(connections),1);    
PlanSetup_ID4(:) = {'Cyber_excel'}; 


data_export = [table(PlanSetup_ID4), connections(:,1:2)];
Actual = zeros([height(connections), 1]) ;

for i = 1:height(connections)
    if ~strcmp(connections{i,3}, 'missing')
        options = length(eval(string(connections{i,3}))); % Mennyi opcio kozul kell kivalasztani
        if options > 1 % Ha nem egyertelmu a nevebol, akkor meg kell keresni a megfelelot
            objective = split(string(connections{i,2}{1}));
            num = double(objective(2));

                if strcmp(objective(1),"V")
                    extension = "Dose";
                else
                    extension = "Volume";
                end
            
                if strcmp(objective(3),"%")
                    extension= extension + "_p";
                elseif strcmp(objective(3),"cGy")
                    extension = extension + "_cGy";
                else
                    extension = extension + "_cm3";
                end
            
            tmp = split(connections{i,3},'.');
            search_word = string(tmp{1}) + "." + string(tmp{2}) + "." + extension;
            array = eval(string(search_word));

            for j = 1:options
                if array(j) == num
                    goal_array = eval(string(connections{i,3}));
                    if strcmp(objective{6},"%") && strcmp(objective{1},"D") %Visszanormalas
                        Actual(i) = goal_array(j) * normal_max / 2235;
                    else
                        Actual(i) = goal_array(j);
                    end
                end
               
            end
        
        else
            tmp = string(split(connections{i,3},'.'));
            if strcmp(tmp(end),"Max")
                objective = split(string(connections{i,2}{1}));
                if strcmp(objective(end),"%")
                Actual(i) = eval(string(connections{i,3}))/2235*100;
                else
                    Actual(i) = eval(string(connections{i,3}));
                end
            else
                Actual(i) = eval(string(connections{i,3}));

            end
        end

    else
        Actual(i) = NaN;
    end

end
Actual = array2table(Actual);

data_export = [data_export,Actual];

data_export=rows2vars(data_export);


location = fullfile( folderPath, "\result" , files(patient).name );

writetable(data_export, location);


%Beleirni a tobbihez
pdf_export = "C:\Users\Gazdi\Documents\SynologyDrive\tanulmanyok\6.BME 2024-2025-2\Szakgyak\pdf_export\excel";
location = fullfile(pdf_export, files(patient).name); 
if isfile(location)
writetable(data_export,location,"WriteMode","append")
%pdf_sum = readtable(location);
%pdf_sum = [pdf_sum; data_export];
%delete(location);
%writetable(pdf_sum,location);
end



end
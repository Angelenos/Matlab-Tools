clear;
clc;

import java.awt.Robot;
import java.awt.event.*;
addpath('Func');
mouse = Robot;
keySet = {'Mule', 'VPA', 'VPB', 'VPC', 'VPD', 'PSA', 'PSB', 'V1', 'V2'};
vehLev = containers.Map(keySet, 0:8);
keySet = {'Module','CDA','Rel_date','Veh_level','MY','Restriction','Spec_Instr'};
property = containers.Map('KeyType','char','ValueType','any');
for i = 1:length(keySet)
    property(keySet{i}) = 0;
end

%% Open log file
flog = fopen(strcat(pwd, '\Log\',datestr(date,'mmddyyyy'),'.log'),'a');
if flog == -1
    error('Can''t open log file');
else
    msg = '-------------------------------------------------------';
    fprintf(flog,'%s\n',msg);
end

%% Determine Screen Resolution
screenSize = get(0, 'MonitorPositions');
numScreen = length(screenSize(:,1));
res1 = screenSize(1,3:4);
res2 = [screenSize(2,3)-screenSize(2,1)+1, screenSize(2,4)];
msg = ['Number of screen: ', num2str(numScreen)];
disp(msg);
fprintf(flog,'%s\n',[datestr(clock,'[HH:MM:SS:FFF] '), msg]);
msg = ['Resolution of Screen 1: ', num2str(res1(1)), 'x' num2str(res1(2))];
disp(msg);
fprintf(flog,'%s\n',[datestr(clock,'[HH:MM:SS:FFF] '), msg]);
msg = ['Resolution of Screen 2: ', num2str(res2(1)), 'x' num2str(res2(2))];
disp(msg);
fprintf(flog,'%s\n',[datestr(clock,'[HH:MM:SS:FFF] '), msg]);
orig = screenSize(2,1:2);

%% User specified options
button = questdlg('Do you want to power off the computer after vehicle summary is updated?',...
    'Power-off Options');
if strcmp(button, 'Cancel')
    msg = 'Program terminated';
    disp(msg);
    fprintf(flog,'%s\n',[datestr(clock,'[HH:MM:SS:FFF] '), msg]);
    fclose(flog);
    return;
else
    msg = ['Automatic Shutdown: ', button];
    disp(msg);
    fprintf(flog,'%s\n',[datestr(clock,'[HH:MM:SS:FFF] '), msg]);
end

%% Get properties of cal release
msg = 'What module to upload: ';
property('Module') = input(msg, 's');
fprintf(flog,'%s\n',[datestr(clock,'[HH:MM:SS:FFF] '), msg, property('Module')]);
msg = 'CDA Version: ';
property('CDA') = input(msg, 's');
fprintf(flog,'%s\n',[datestr(clock,'[HH:MM:SS:FFF] '), msg, property('CDA')]);
msg = 'Release date (mm/dd/yy Leave blank for today): ';
property('Rel_Date') = input(msg, 's');
if isempty(property('Rel_Date'))
    property('Rel_Date') = datestr(date,'mm/dd/yyyy');
end
fprintf(flog,'%s\n',[datestr(clock,'[HH:MM:SS:FFF] '), msg, property('Rel_Date')]);
Veh_level = 0;
msg = 'Release level: ';
while ~vehLev.isKey(Veh_level)
    Veh_level = input(msg, 's');
end
fprintf(flog,'%s\n',[datestr(clock,'[HH:MM:SS:FFF] '), msg, Veh_level]);
property('Veh_level') = vehLev(Veh_level);
msg = 'Model Year: ';
property('MY') = input(msg);
fprintf(flog,'%s\n',[datestr(clock,'[HH:MM:SS:FFF] '), msg, ...
    num2str(property('MY'))]);

%% Load Part Number List
filename = dir('PN List\*.xlsx');
if length(filename) < 1
    msg = 'ERROR: No Part Number List Found';
    fprintf(flog,'%s\n',[datestr(clock,'[HH:MM:SS:FFF] '), msg]);
    fclose(flog);
    error(msg);
end
filename = strcat(pwd,'\PN List\',filename(1).name);
[~,~,pn_list_raw] = xlsread(filename);
pn_list_raw = pn_list_raw(2:end,:);
pn_list = pn_list_raw(:,3);
if(size(pn_list_raw,2)>13)
    has_spec_instr = 1;
else
    has_spec_instr = 0;
end
for i = 1:length(pn_list_raw(:,3))
    pn_list(i) = regexp(pn_list_raw(i,3),'(?<=\[)\w+(?=\])','match');
    pn_list{i} = pn_list{i}{1};
end
crit = strcat({'(?<='},pn_list(:),{'[\_| ])[\S\s]+'});
for i = 1:length(pn_list_raw(:,3))
    pn_list(i,2) = regexp(pn_list_raw(i,3),crit(i),'match');
    if(~isempty(pn_list{i,2}))
        pn_list{i,2} = pn_list{i,2}{1};
    end
end
if(has_spec_instr)
    pn_list(:,3) = pn_list_raw(:,14);
end
num_efd_list = length(pn_list(:,1));
msg = ['Part Number List Loaded. Found ', num2str(num_efd_list),...
    ' efd Files in Total'];
fprintf(flog,'%s\n',[datestr(clock,'[HH:MM:SS:FFF] '), msg]);
disp(msg);

%% Get local folder of efd files
efd_folder = uigetdir('C:\work\OBD reliability\Calibration');
if efd_folder == 0
    msg = 'ERROR: No folder specified';
    fprintf(flog,'%s\n',[datestr(clock,'[HH:MM:SS:FFF] '), msg]);
    fclose(flog);
    error(msg);
else
    msg = ['Local Folder: ', efd_folder];
    fprintf(flog,'%s\n',[datestr(clock,'[HH:MM:SS:FFF] '), msg]);
    disp(msg);
end
efdFiles = dir(strcat(efd_folder,'\*.efd'));
num_efd_loc = length(efdFiles);
if num_efd_loc ~= num_efd_list
    msg = 'ERROR: Numbers of Files from PowerCal don''t match local folders';
    fprintf(flog,'%s\n',[datestr(clock,'[HH:MM:SS:FFF] '), msg]);
    fclose(flog);
    error(msg);
else
    msg = ['Found ', num2str(num_efd_loc), ' efd files from local folders'];
    fprintf(flog,'%s\n',[datestr(clock,'[HH:MM:SS:FFF] '), msg]);
    disp(msg);
end

%% Get prepared for upload
uiwait(msgbox('Make sure the Chrome is opened and Sharepoint website is loaded',...
    'Before uploading','warn'));

get_page_layout = 0;
pause(2);

%% Uploading Files
tic;
for i = 1:num_efd_loc
    % Choose efd file and upload
    msg = ['Processing ', num2str(i), '/', num2str(num_efd_loc), ' files: ',...
        pn_list{i,1}];
    fprintf(flog,'%s\n',[datestr(clock,'[HH:MM:SS:FFF] '), msg]);
    disp(msg);
    if(has_spec_instr)
        property('Spec_Instr') = pn_list{i,3};
        if(isempty(property('Spec_Instr')))
            property('Spec_Instr') = 0;
        end
    else
        property('Spec_Instr') = 0;
    end
    property('Restriction') = pn_list{i,2};
    mouse_move(mouse, orig, res2, [12, 4.5]);
    mouse_click(mouse);
    mouse.delay(500)
    mouse_move(mouse, orig, res2, [2.1, 2.2]);
    mouse_click(mouse);
    mouse.delay(500)
    mouse_move(mouse, orig, res2, [1.4, 16]);
    mouse_click(mouse);
    clipboard('copy',efd_folder);
    mouse_paste(mouse);
    mouse.delay(300)
    mouse_move(mouse, orig, res2, [5, 1.1]);
    mouse_click(mouse);
    mouse.delay(300)
    clipboard('copy',strcat('*',pn_list{i,1},'*'));
    mouse_paste(mouse);
    mouse.delay(300)
    mouse_move(mouse, orig, res2, [3, 5.3]);
    mouse_click(mouse);
    mouse.delay(200);
    mouse_move(mouse, orig, res2, [1.15, 1.05]);
    mouse_click(mouse);
    mouse.delay(500);
    mouse_move(mouse, orig, res2, [2, 1.4]);
    mouse_click(mouse);
    wait_next(mouse, orig, res2, [2.7, 1.1], 750);
    if ~get_page_layout
        mouse_move(mouse, orig, res2, [3, 1.5]);
        mouse.delay(500);
        mouse_frame(mouse, orig, res2);
        mouse.delay(1800);
        mouse_all(mouse);
        mouse.delay(300);
        mouse_copy(mouse);
        mouse.delay(300);
        mouse_close(mouse);
        page_src = clipboard('paste');
        field_arr = regexp(page_src,'(?<=FieldName=")[\w+\s]+\w+','match');
        get_page_layout = 1;
    end
    % Edit properties
    property_edit(mouse, field_arr, property, orig, res2);
end

timestamp = toc;
msg = ['Total uploading time is ', num2str(timestamp), ' seconds'];
fprintf(flog,'%s\n',[datestr(clock,'[HH:MM:SS:FFF] '), msg]);
disp(msg);
fclose(flog);

if strcmp(button,'Yes')
    system('shutdown -s -t 60'); % Power off
end

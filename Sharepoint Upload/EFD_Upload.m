clear;
clc;

import java.awt.Robot;
import java.awt.event.*;
mouse = Robot;
keySet = {'Mule', 'VPA', 'VPB', 'VPC', 'VPD', 'PSA', 'PSB', 'V1', 'V2'};
vehLev = containers.Map(keySet, 0:8);

%% Open log file
flog = fopen(strcat(pwd, '\Log\',datestr(date,'mmddyyyy'),'.log'),'a');
if flog == -1
    error('Can''t open log file');
else
    msg = '-------------------------------------------------------';
    fprintf(flog,'%s\n',msg);
end

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

%% Get number and resolution of screen
msg = 'What module to upload: ';
Module = input(msg, 's');
fprintf(flog,'%s\n',[datestr(clock,'[HH:MM:SS:FFF] '), msg, Module]);
msg = 'CDA Version: ';
CDA_ver = input(msg, 's');
fprintf(flog,'%s\n',[datestr(clock,'[HH:MM:SS:FFF] '), msg, CDA_ver]);
msg = 'Release date (mm/dd/yy Leave blank for today): ';
Rel_Date = input(msg, 's');
if isempty(Rel_Date)
    Rel_Date = datestr(date,'mm/dd/yyyy');
end
fprintf(flog,'%s\n',[datestr(clock,'[HH:MM:SS:FFF] '), msg, Rel_Date]);
Veh_level = 0;
msg = 'Release level: ';
while ~vehLev.isKey(Veh_level)
    Veh_level = input(msg, 's');
end
fprintf(flog,'%s\n',[datestr(clock,'[HH:MM:SS:FFF] '), msg, Veh_level]);
Veh_level = vehLev(Veh_level);
msg = 'Model Year: ';
MY = input(msg);
fprintf(flog,'%s\n',[datestr(clock,'[HH:MM:SS:FFF] '), msg, num2str(MY)]);
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
for i = 1:length(pn_list_raw(:,3))
    pn_list(i) = regexp(pn_list_raw(i,3),'(?<=\[)\w+(?=\])','match');
    pn_list{i} = pn_list{i}{1};
end
crit = strcat({'(?<='},pn_list(:),{'[\_| ])[\S\s]+'});
for i = 1:length(pn_list_raw(:,3))
    pn_list(i,2) = regexp(pn_list_raw(i,3),crit(i),'match');
    pn_list{i,2} = pn_list{i,2}{1};
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
orig = screenSize(2,1:2);
pause(2);

%% Uploading Files
tic;
for i = 1:num_efd_loc
    % Choose efd file and upload
    msg = ['Processing ', num2str(i), '/', num2str(num_efd_loc), ' files: ',...
        pn_list{i,1}];
    fprintf(flog,'%s\n',[datestr(clock,'[HH:MM:SS:FFF] '), msg]);
    disp(msg);
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
    % Edit properties
    mouse.keyPress(KeyEvent.VK_TAB);
    mouse.keyRelease(KeyEvent.VK_TAB);
    mouse.delay(200);
    mouse.keyPress(KeyEvent.VK_TAB);
    mouse.keyRelease(KeyEvent.VK_TAB);
    clipboard('copy', Module);
    mouse_paste(mouse);
    mouse.delay(200);
    mouse.keyPress(KeyEvent.VK_TAB);
    mouse.keyRelease(KeyEvent.VK_TAB);
    clipboard('copy', CDA_ver);
    mouse_paste(mouse);
    mouse.delay(200);
    mouse.keyPress(KeyEvent.VK_TAB);
    mouse.keyRelease(KeyEvent.VK_TAB);
    clipboard('copy', Rel_Date);
    mouse_paste(mouse);
    mouse.delay(200);
    mouse.keyPress(KeyEvent.VK_TAB);
    mouse.keyRelease(KeyEvent.VK_TAB);
    mouse.delay(200);
    mouse.keyPress(KeyEvent.VK_TAB);
    mouse.keyRelease(KeyEvent.VK_TAB);
    for j = 1:Veh_level
        mouse.keyPress(KeyEvent.VK_DOWN);
        mouse.keyRelease(KeyEvent.VK_DOWN);
        mouse.delay(200);
    end
    mouse.keyPress(KeyEvent.VK_TAB);
    mouse.keyRelease(KeyEvent.VK_TAB);
    mouse.delay(200);
    for j = 1: MY-2011
        mouse.keyPress(KeyEvent.VK_TAB);
        mouse.keyRelease(KeyEvent.VK_TAB);
        mouse.delay(200);
    end
    mouse.keyPress(KeyEvent.VK_SPACE);
    mouse.keyRelease(KeyEvent.VK_SPACE);
    mouse.delay(200);
    for j = 1: 2030-MY
        mouse.keyPress(KeyEvent.VK_TAB);
        mouse.keyRelease(KeyEvent.VK_TAB);
        mouse.delay(200);
    end
    mouse.keyPress(KeyEvent.VK_TAB);
    mouse.keyRelease(KeyEvent.VK_TAB);
    mouse.delay(200);
    mouse.keyPress(KeyEvent.VK_TAB);
    mouse.keyRelease(KeyEvent.VK_TAB);
    mouse.delay(200);
    clipboard('copy',pn_list{i,2});
    mouse_all(mouse)
    mouse.delay(300);
    mouse_paste(mouse);
    mouse.delay(200);
    mouse.keyPress(KeyEvent.VK_TAB);
    mouse.keyRelease(KeyEvent.VK_TAB);
    mouse.delay(200);
    clipboard('copy','NA');
    mouse_all(mouse)
    mouse.delay(300);
    mouse_paste(mouse);
    mouse.delay(200);
    for j = 1:4
        mouse.keyPress(KeyEvent.VK_TAB);
        mouse.keyRelease(KeyEvent.VK_TAB);
        mouse.delay(200);
    end
    mouse.keyPress(KeyEvent.VK_ENTER);
    mouse.keyRelease(KeyEvent.VK_ENTER);
    wait_next(mouse, orig, res2, [1.1, 1.1], 750);
    mouse.delay(1000);
end

timestamp = toc;
msg = ['Total uploading time is ', num2str(timestamp), ' seconds'];
fprintf(flog,'%s\n',[datestr(clock,'[HH:MM:SS:FFF] '), msg]);
disp(msg);
fclose(flog);

if strcmp(button,'Yes')
    system('shutdown -s -t 60'); % Power off
end

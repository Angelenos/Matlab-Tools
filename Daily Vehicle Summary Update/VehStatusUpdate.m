clear;
clc;

import java.awt.Robot;
import java.awt.event.*;

mouse = Robot;
screenSize = get(0, 'MonitorPositions');
if(screenSize(3) == 1920)
    offset = 1920;
else
    offset = 0;
end

saveLoc = 'C:\work\OBD reliability\Issue Tracking\Vehicle Status Summary';

%% Open log file
flog = fopen(strcat(pwd, '\Log\',datestr(date,'mmddyyyy'),'.log'),'a');
if flog == -1
    error('Can''t open log file');
else
    msg = '-------------------------------------------------------';
    fprintf(flog,'%s\n',msg);
end

%% Check if needs update
files = dir(strcat(saveLoc,'\*.csv'));
filenames = {files.name};
if sum(~cellfun('isempty',strfind(filenames,datestr(now,'mmdd')))) > 0
    uiwait(msgbox('Files have been updated today',...
        'Update suspended','error'));
    return;
end

%% User-specified options

button = questdlg('Do you want to power off the computer after vehicle summary is updated?',...
    'Power-off Options');
if strcmp(button, 'Cancel')
    msg = 'Program terminated';
    disp(msg);
    fprintf(flog,'%s\n',[datestr(clock,'[HH:MM:SS:FFF] '), msg]);
    fclose(flog);
    return;
elseif strcmp(button, 'Yes')
    msg = ['Automatic Shutdown: ', button];
    disp(msg);
    fprintf(flog,'%s\n',[datestr(clock,'[HH:MM:SS:FFF] '), msg]);
    button2 = questdlg('Do you want to power off the computer when error presented?',...
        'Power-off Options','Yes','No','Yes');
    if isempty(button2)
        button2 = 'No';
    end
    msg = ['Shutdown ignoring error presented: ', button];
    disp(msg);
    fprintf(flog,'%s\n',[datestr(clock,'[HH:MM:SS:FFF] '), msg]);
end
uiwait(msgbox('Make sure the Chrome is opened and ControlTec website is loaded',...
    'One more thing...','warn'));

pause(5);

%% Vehicle List Download
mouse.mouseMove(2050 - offset, 360);
mouse.delay(500)
mouse.mousePress(InputEvent.BUTTON1_MASK)
mouse.delay(500)
mouse.mouseRelease(InputEvent.BUTTON1_MASK)
mouse.delay(5000)

% Download
tic
while 1
    colortemp = mouse.getPixelColor(3204 - offset, 390);
    if colortemp.getRed + colortemp.getRed + colortemp.getRed >= 250
        break;
    end
end
timestamp = toc;
msg = ['Page loading time: ', num2str(timestamp), ' seconds'];
disp(msg);
fprintf(flog,'%s\n',[datestr(clock,'[HH:MM:SS:FFF] '), msg]);
download_mouse(1)
mouse.delay(1000)
msg = 'Vehicle list has been downloaded';
disp(msg);
fprintf(flog,'%s\n',[datestr(clock,'[HH:MM:SS:FFF] '), msg]);

%% Vehicle Property Download
mouse.mouseMove(2275 - offset, 360);
mouse.delay(500)
mouse.mousePress(InputEvent.BUTTON1_MASK)
mouse.delay(500)
mouse.mouseRelease(InputEvent.BUTTON1_MASK)
mouse.delay(5000)

% Download
tic
while 1
    colortemp = mouse.getPixelColor(3206 - offset, 390);
    if colortemp.getRed + colortemp.getRed + colortemp.getRed >= 250
        break;
    end
end
timestamp = toc;
msg = ['Page loading time: ', num2str(timestamp), ' seconds'];
disp(msg);
fprintf(flog,'%s\n',[datestr(clock,'[HH:MM:SS:FFF] '), msg]);
download_mouse(2)
mouse.delay(1000)
msg = 'Vehicle property summary has been downloaded';
disp(msg);
fprintf(flog,'%s\n',[datestr(clock,'[HH:MM:SS:FFF] '), msg]);

%% Vehicle Software Download
mouse.mouseMove(2355 - offset, 360);
mouse.delay(500)
mouse.mousePress(InputEvent.BUTTON1_MASK)
mouse.delay(500)
mouse.mouseRelease(InputEvent.BUTTON1_MASK)
mouse.delay(5000)

% Download
tic
while 1
    colortemp = mouse.getPixelColor(3207 - offset, 390);
    if colortemp.getRed + colortemp.getRed + colortemp.getRed >= 250
        break;
    end
end
timestamp = toc;
msg = ['Page loading time: ', num2str(timestamp), ' seconds'];
disp(msg);
fprintf(flog,'%s\n',[datestr(clock,'[HH:MM:SS:FFF] '), msg]);
download_mouse(3)
mouse.delay(1000)
msg = 'Vehicle software summary has been downloaded';
disp(msg);
fprintf(flog,'%s\n',[datestr(clock,'[HH:MM:SS:FFF] '), msg]);

%% Downloaded File Name Check
[n_file, file_name_err] = name_check();
msg = ['Number of files with corrupted name: ', num2str(n_file)];
disp(msg);
fprintf(flog,'%s\n',[datestr(clock,'[HH:MM:SS:FFF] '), msg]);
for i = 1:n_file
    msg = ['   - File name being corrected: ', file_name_err{i}];
    disp(msg);
    fprintf(flog,'%s\n',[datestr(clock,'[HH:MM:SS:FFF] '), msg]);
end

%% Update Summary Sheet
ExcelApp = actxserver('Excel.Application');
ExcelApp.Visible = 1;
VehSummarySheet = ExcelApp.Workbooks.Open(strcat(saveLoc,...
    '\Vehicle Status Summary.xlsm'));
msg = ['Running macro to update vehicle status summary'];
disp(msg);
fprintf(flog,'%s\n',[datestr(clock,'[HH:MM:SS:FFF] '), msg]);
try
    ExcelApp.Run('Sheet1.Veh_Status_Auto_Fill');
catch MacError
    if (strfind(MacError.message,'Error: Object returned error code:'))
        VehSummarySheet.Save;
        ExcelApp.Quit;
        ExcelApp.delete;
        copyfile (strcat(saveLoc,'\old\Vehicle Status Summary_',...
            datestr(now,'mmdd'),'.xlsm'), strcat(saveLoc,...
            '\Vehicle Status Summary.xlsm'),'f');
        msg = ['Error in Macro Running: ',MacError.message];
        fprintf(flog,'%s\n',[datestr(clock,'[HH:MM:SS:FFF] '), msg]);
        msgbox({'Error in Macro Running',MacError.message},...
        'Error','error');
        if(strcmp(button2, 'Yes'))
            system('shutdown -s -t 60'); % Power off
        end
        fclose(flog);
        return;
    end
end
pause(10);
VehSummarySheet.Save;
ExcelApp.Quit;
ExcelApp.delete;

%% Power-off if Excel is shut down correctly
pause(5);
msg = ['Vehicle status summary has been updated'];
disp(msg);
fprintf(flog,'%s\n',[datestr(clock,'[HH:MM:SS:FFF] '), msg]);
if strcmp(button,'Yes')
    msg = ['Computer will be automatically shutdown in 60 seconds'];
    disp(msg);
    fprintf(flog,'%s\n',[datestr(clock,'[HH:MM:SS:FFF] '), msg]);
    system('shutdown -s -t 60'); % Power off
end
fclose(flog);
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
    disp('Program terminated.');
    return;
elseif strcmp(button, 'Yes')
    button2 = questdlg('Do you want to power off the computer when error presented?',...
        'Power-off Options','Yes','No','Yes');
    if isempty(button2)
        button2 = 'No';
    end
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
disp(['Page loading time: ', num2str(timestamp), ' seconds']);
download_mouse(1)
mouse.delay(1000)

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
disp(['Page loading time: ', num2str(timestamp), ' seconds']);
download_mouse(2)
mouse.delay(1000)

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
disp(['Page loading time: ', num2str(timestamp), ' seconds']);
download_mouse(3)
mouse.delay(1000)

%% Downloaded File Name Check
n_file = name_check();
disp(['Number of files with corrupted name: ', num2str(n_file)]);

%% Update Summary Sheet
ExcelApp = actxserver('Excel.Application');
ExcelApp.Visible = 1;
VehSummarySheet = ExcelApp.Workbooks.Open(strcat(saveLoc,...
    '\Vehicle Status Summary.xlsm'));
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
        msgbox({'Error in Macro Running',MacError.message},...
        'Error','error');
        if(strcmp(button2, 'Yes'))
            system('shutdown -s -t 60'); % Power off
        end
        return;
    end
end
pause(10);
VehSummarySheet.Save;
ExcelApp.Quit;
ExcelApp.delete;

%% Power-off if Excel is shut down correctly
pause(5);

[~,task] = system('tasklist/fi "imagename eq Excel.exe"');
if length(task) < 100 && strcmp(button,'Yes')
    system('shutdown -s -t 60'); % Power off
end
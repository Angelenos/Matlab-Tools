function download_mouse(i)
    import java.awt.Robot;
    import java.awt.event.*;
    mouse = Robot;
    
    saveLoc = 'C:\work\OBD reliability\Issue Tracking\Vehicle Status Summary';
    vehdetName = strcat('CONTROLTEC Vehicle Detail_', datestr(now,'mmdd'),'.csv');
    vehproName = strcat('CONTROLTEC Vehicle Properties_', datestr(now,'mmdd'),'.csv');
    vehsoftName = strcat('CONTROLTEC Vehicle Software Detail_', datestr(now,'mmdd'),'.csv');
    filenames = {vehdetName, vehproName, vehsoftName};
    
    screenSize = get(0, 'MonitorPositions');
    if(screenSize(3) == 1920)
        offset = 1920;
    else
        offset = 0;
    end
    
    mouse.mouseMove(3205 - offset, 390);
    mouse.mousePress(InputEvent.BUTTON2_MASK)
    mouse.delay(500)
    mouse.mouseRelease(InputEvent.BUTTON2_MASK)
    
    tic
    while 1
        colortemp = mouse.getPixelColor(3205 - offset, 390);
        if colortemp.getRed + colortemp.getRed + colortemp.getRed >= 750
            break;
        end
    end
    
    timestamp = toc;
    disp(['CSV generation time: ', num2str(timestamp), ' seconds']);
    
    mouse.delay(1000)
    mouse.mouseMove(2900 - offset, 55);
    mouse.delay(500)
    mouse.mousePress(InputEvent.BUTTON1_MASK);
    mouse.mouseRelease(InputEvent.BUTTON1_MASK);
    clipboard('copy',saveLoc);
    mouse.delay(500)
    mouse.keyPress(KeyEvent.VK_CONTROL);
    mouse.delay(500)
    mouse.keyPress(KeyEvent.VK_V);
    mouse.keyRelease(KeyEvent.VK_V);
    mouse.keyRelease(KeyEvent.VK_CONTROL);
    mouse.delay(500)
    mouse.keyPress(KeyEvent.VK_ENTER);
    mouse.keyRelease(KeyEvent.VK_ENTER);
    mouse.mouseMove(2900 - offset, 770);
    mouse.delay(1000)
    mouse.mousePress(InputEvent.BUTTON1_MASK);
    mouse.mouseRelease(InputEvent.BUTTON1_MASK);
    clipboard('copy',filenames{i});
    mouse.delay(500)
    mouse.keyPress(KeyEvent.VK_CONTROL);
    mouse.delay(500)
    mouse.keyPress(KeyEvent.VK_V);
    mouse.keyRelease(KeyEvent.VK_V);
    mouse.keyRelease(KeyEvent.VK_CONTROL);
    mouse.delay(500)
    mouse.keyPress(KeyEvent.VK_ENTER);
    mouse.keyRelease(KeyEvent.VK_ENTER);
end
function property_edit(mouse, field_arr, property, orig, res2)
    import java.awt.event.*;
    mouse_move(mouse, orig, res2, [1.8, 2.2]);
    mouse_click(mouse);
    for i = 2:length(field_arr)
        mouse.keyPress(KeyEvent.VK_TAB);
        mouse.keyRelease(KeyEvent.VK_TAB);
        mouse.delay(200);
        field = field_arr{i};
        mouse.delay(200);
        switch field
            case 'CDA Software Version'
                clipboard('copy', property('CDA'));
                mouse_paste(mouse);
            case 'Restrictions'                              
                mouse_all(mouse)
                clipboard('copy', property('Restriction'));
                mouse_paste(mouse);  
            case 'Special Instructions'
                mouse_all(mouse)
                if property('Spec_Instr')==0
                    clipboard('copy', 'NA');
                else
                    clipboard('copy', property('Spec_Instr'));
                end
                mouse_paste(mouse);
            case 'Module'
                clipboard('copy', property('Module'));
                mouse_paste(mouse);
            case 'Release Date'
                clipboard('copy', property('Rel_Date'));
                mouse_paste(mouse);
                mouse.keyPress(KeyEvent.VK_TAB);
                mouse.keyRelease(KeyEvent.VK_TAB);
            case 'Model Year'
                for j = 1: property('MY')-2011
                    mouse.keyPress(KeyEvent.VK_TAB);
                    mouse.keyRelease(KeyEvent.VK_TAB);
                    mouse.delay(200);
                end
                mouse.keyPress(KeyEvent.VK_SPACE);
                mouse.keyRelease(KeyEvent.VK_SPACE);
                mouse.delay(200);
                for j = 1: 2030-property('MY')
                    mouse.keyPress(KeyEvent.VK_TAB);
                    mouse.keyRelease(KeyEvent.VK_TAB);
                    mouse.delay(200);
                end
            case 'Vehicle Level'
                for j = 1:property('Veh_level')
                    mouse.keyPress(KeyEvent.VK_DOWN);
                    mouse.keyRelease(KeyEvent.VK_DOWN);
                    mouse.delay(200);
                end
                mouse.keyPress(KeyEvent.VK_TAB);
                mouse.keyRelease(KeyEvent.VK_TAB);
            otherwise
                mouse.delay(200);
        end
    end
    mouse.keyPress(KeyEvent.VK_TAB);
    mouse.keyRelease(KeyEvent.VK_TAB);
    mouse.delay(200);
    mouse.keyPress(KeyEvent.VK_TAB);
    mouse.keyRelease(KeyEvent.VK_TAB);
    mouse.delay(200);
    mouse.keyPress(KeyEvent.VK_ENTER);
    mouse.keyRelease(KeyEvent.VK_ENTER);
    wait_next(mouse, orig, res2, [1.1, 1.1], 750);
    mouse.delay(1000);
end
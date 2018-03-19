function mouse_click(mouse, key)
    import java.awt.event.*;
    if nargin == 1
        key_mask = InputEvent.BUTTON1_MASK;
    else
        if key == 1
            key_mask = InputEvent.BUTTON3_MASK;
        else
            key_mask = InputEvent.BUTTON2_MASK;
        end
    end
    mouse.mousePress(key_mask)
    mouse.delay(200)
    mouse.mouseRelease(key_mask)
    mouse.delay(300)
end
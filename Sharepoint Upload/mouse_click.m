function mouse_click(mouse)
    import java.awt.event.*;
    mouse.mousePress(InputEvent.BUTTON1_MASK)
    mouse.delay(200)
    mouse.mouseRelease(InputEvent.BUTTON1_MASK)
    mouse.delay(300)
end
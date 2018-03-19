function mouse_copy(mouse)
    import java.awt.event.*;
    mouse.keyPress(KeyEvent.VK_CONTROL);
    mouse.delay(300)
    mouse.keyPress(KeyEvent.VK_C);
    mouse.delay(200)
    mouse.keyRelease(KeyEvent.VK_C);
    mouse.keyRelease(KeyEvent.VK_CONTROL);
end
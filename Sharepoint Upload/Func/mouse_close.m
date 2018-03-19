function mouse_close(mouse)
    import java.awt.event.*;
    mouse.keyPress(KeyEvent.VK_CONTROL);
    mouse.delay(300)
    mouse.keyPress(KeyEvent.VK_W);
    mouse.delay(200)
    mouse.keyRelease(KeyEvent.VK_W);
    mouse.keyRelease(KeyEvent.VK_CONTROL);
    mouse.delay(200)
end
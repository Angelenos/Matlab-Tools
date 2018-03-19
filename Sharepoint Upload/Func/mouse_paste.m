function mouse_paste(mouse)
    import java.awt.event.*;
    mouse.keyPress(KeyEvent.VK_CONTROL);
    mouse.delay(300)
    mouse.keyPress(KeyEvent.VK_V);
    mouse.delay(200)
    mouse.keyRelease(KeyEvent.VK_V);
    mouse.keyRelease(KeyEvent.VK_CONTROL);
    mouse.delay(200)
    mouse.keyPress(KeyEvent.VK_ENTER);
    mouse.keyRelease(KeyEvent.VK_ENTER);
end
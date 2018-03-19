function mouse_all(mouse)
    import java.awt.event.*;
    mouse.keyPress(KeyEvent.VK_CONTROL);
    mouse.delay(300)
    mouse.keyPress(KeyEvent.VK_A);
    mouse.delay(200)
    mouse.keyRelease(KeyEvent.VK_A);
    mouse.keyRelease(KeyEvent.VK_CONTROL);
    mouse.delay(200)
end
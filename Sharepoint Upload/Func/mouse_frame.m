function mouse_frame(mouse, orig, res2)
    import java.awt.event.*;
    mouse_move(mouse, orig, res2, [3, 2.2]);
    mouse_click(mouse, 1);
    mouse_move(mouse, orig, res2, [2.5, 1.3]);
    mouse_click(mouse);
end
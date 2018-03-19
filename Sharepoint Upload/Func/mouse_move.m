function mouse_move(mouse, orig, res, coord)
    x = coord(1);
    y = coord(2);
    mouse.mouseMove(orig(1)+(res(1)/x),orig(2)+(res(2)/y));
    mouse.delay(300)
end
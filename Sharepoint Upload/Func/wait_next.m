function wait_next(mouse, orig, res, coord, thres)
    x = coord(1);
    y = coord(2);
    while 1
        colortemp = mouse.getPixelColor(orig(1)+(res(1)/x),orig(2)+(res(2)/y));
        if colortemp.getRed + colortemp.getRed + colortemp.getRed >= thres
            break;
        end
    end
end